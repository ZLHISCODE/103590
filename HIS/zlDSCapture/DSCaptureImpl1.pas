{*******************************************************************************
视频采集COM对象实现单元
创建人：TJH
创建日前：2009-11-3

描述：...

DirectShow格式说明：

MEDIATYPE_Video;    ／／纯视频(数字) 
MEDIATYPE_Audio;    ／／纯音频(数字)
MEDIATYPE_AnalogVideo;    ／／模拟视频，一般是视频采集卡输入的数据类型 
MEDIATYPE_AnalogAudio;    ／／模拟音频，一般是声卡采集输入的数据类型 
MEDIATYPE_Text；    ／／文字 
MEDIATYPE_Midi;    ／／MIDI音乐 
MEDIATYPE_STREAM;  //字节流,如(Pull模式)文件源的输出数据类型 
MEDIATYPE_Interleaved;    ／／数码摄像机输入的DV数据类型 
MEDIATYPE_MPEG1SystemStream;    ／／MPEG1的系统流 
MEDIATYPE_MPEG2_PACK;    ／／MPEG2的数据包 
MEDIATYPE_MPEG2_PES;    ／／MPEG2分组数据 
MEDIATYPE_DVD_ENCRYPTED_PACK;    ／／DVD播放用到的媒体类型 
MEDIATYPE_DVD_NAVIGATION;


媒体类型主要用3部分来描述：majortype(主类型)、subtype(辅助说明类型)和formattype(格式细节类型)。
这3部分各自用一个GUID来标识。它们的作用分别是：majortype定性地描述媒体类型，
如指定这是一个视频 (MEDIATYPE_Video)、音频(MEDIATYPE_Audio)或者字节流(MEDIATYPE_Stream)等；

subtype辅助说明majortype，指明具体是哪种格式，例如，若majortype是视频，
subtype可以进一步指明这是UYVY(MEDIASUBTYPE_UYVY)、YUY2(MEDIASUBTYPE_YUY2)、
RGB24(MEDIASUBTYPE_RGB24)还是RGB32(MEDIASUBTYPE_RGB32)等，若majortype是音频，
subtype可以进一步指明这是PCM格式(MEDIASUBTYPE_PCM)还是AC3格式(MEDIASUBTYPE_DOLBY_AC3)等;

formattype指定了一种进一步描述格式细节的数据结构类型，格式细节描述的内容主要包括视频图像的大小、
帧率，或音频的采样频率、量化精度等参数，这个描述格式细节的数据块指针保存在pbFormat成员中。


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
    _CaptureParameter: TCaptureParameter;  //采集参数
    _frmCaptureParameterCfg: TfrmCapParameterCfg;

    _CaptureParameterCfgFileName: WideString; //采集参数的配置文件名称

    _IsPreviewState: Boolean;       //是否为预览状态
    _IsCaptureVideo: Boolean;     //是否视频采集
    _TempCaptureVideoFile: WideString; //临时采集文件名称

    _IsStretch: Boolean;             //自动填充窗口大小
    _IsAdjustWindowSize: Boolean;    //自动调整窗口大小
    _IsFit: Boolean;                 //自动适应窗口大小

    _GraphManager: ICaptureGraphBuilder2; //管理GraphBuilder中的所有FILTER
    _CapSourceFilter: IBaseFilter;
    _EncoderFilter: IBaseFilter;      //视频编码器
    _AviMultiplexer: IBaseFilter;     //多路服用接口
    _AviWriter: IBaseFilter;          //文件写入接口
    _SmartTee: IBaseFilter;           //数据分流接口
    _SmartTee1: IBaseFilter;          //针对天朗采集卡，PX1000E，这种采集卡需要连接两个SmartTee1的preview才能是显示的视频不卡
    _ColorSpace: IBaseFilter;         //颜色转换Filter
    //_MjpegDescompress: IBaseFilter;   //mjpeg压缩接口

    _HideCfgItem: Integer;            //需要隐藏的配置项

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


    //初始化DSHOW的相关配置
    procedure InitDShow();
    //反初始化DSHOW的相关配置
    procedure ReInitDShow();


    //参数修改事件
    procedure ParameterChange(const capParameter: TCaptureParameter; const needCaptureSample: Boolean);
    //选项中的vfw配置调用事件
    procedure vfwConfigCallEvent(const operVfwConfigType: TVfwConfigType;
      const parentHandle: Integer; out errMsg: WideString);

    //设置视频质量
    procedure ConfigVideoQuality(filter: IBaseFilter; captureParameter: TCaptureParameter);
    //设置视频制式
    procedure ConfigVideoAnalog(filter: IBaseFilter; captureParameter: TCaptureParameter);
    //设置视频格式
    procedure ConfigVideoFormat(filter: IBaseFilter; captureParameter: TCaptureParameter);


    //根据设置调整窗口大小
    procedure AdjustWindowSize();
    //采集图像到BMP对象
    function CaptureImageToBmpObj(): TBitmap;
    //显示VFW配置对话框
    procedure ShowVfwConfigDialog(const dialogType: Integer; const parentHandle: Integer);
    //显示VFW压缩编码设置
    procedure ShowVfwCompressCfgDialog(const dialogType: Integer; const parentHandle: Integer);

    //预览
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

    //property   保留IsStretch和IsFit属性主要是为了和以前公开的接口兼容
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
    
    //释放资源
    procedure FreeRes; safecall;
    //开始预览
    function StartPreview: WideString; safecall;
    //停止预览
    function StopPreview: WideString; safecall;
    //采集BMP图像到文件
    function CaptureBmpImageToFile(const fileName: WideString): WideString; safecall;
    //采集JPG图像到文件
    function CaptureJpgImageToFile(const fileName: WideString; compressRate: SYSINT): WideString; safecall;
    //开始视频采集
    function StartCaptureVideo(const fileName: WideString): WideString; safecall;
    //停止视频采集
    function StopCaptureVideo(out videoFile: WideString): WideString; safecall;
    //采集参数设置  -- 不需要ParentHandle的值
    function ShowCaptureParameterCfgDialog(parentHandle: Integer): WideString; safecall;
    
    //显示采集源filter配置  -- 需要驱动支持
    function ShowVideoCaptureFilterCfg(parentHandle: Integer): WideString; safecall;
    //显示视频编码器属性  -- 需要驱动支持
    function ShowVideoEncoderFilterCfg(parentHandle: Integer): WideString; safecall;
    //显示采集端口配置  -- 需要驱动支持
    function ShowVideoCapturePinCfg(parentHandle: Integer): WideString; safecall;

    //显示VFW显示方式设置  -- 需要驱动支持
    function ShowVfwVideoDisplayCfg(parentHandle: Integer): WideString; safecall;
    //显示VFW视频格式配置  -- 需要驱动支持
    function ShowVfwVideoFormatCfg(parentHandle: Integer): WideString; safecall;
    //显示视频源配置  -- 需要驱动支持
    function ShowVfwVideoSourceCfg(parentHandle: Integer): WideString; safecall;
    
    //从配置文件读取采集参数
    function ReadParameterFromFile: WideString; safecall;
    //刷新窗口
    function RefreshWindow: WideString; safecall;
    //退出全屏
    function QuitFullScreen: WideString; safecall;
    //全屏显示
    function ShowFullScreen(parentHandle, monitorIndex: Integer): WideString; safecall;
    //更新视频质量
    function UpdateVideoQuailty: WideString; safecall;
    //保存采集参数到配置文件
    function SaveParameterToFile: WideString; safecall;
    //取得视频采集参数
    function GetCaptureParameter(var parameter: TCaptureParameter): WideString; safecall;
    //设置采集参数
    function SetCaptureParameter(var parameter: TCaptureParameter): WideString; safecall;
    //重新预览
    function RePreview: WideString; safecall;
    function CaptureImgToClipBoard: WideString; safecall;
    //显示vfw压缩设置
    function ShowVfwCompressCfg(parentHandle: Integer): WideString; safecall;
    //显示videoCrossbar设置
    function ShowVideoCrossbarCfg(parentHandle: Integer): WideString; safecall;
    //采集bmp图像
    function CaptureBmpImage: IPictureDisp; safecall;
    //采集jpg图像(当转换成IPictureDisp后，数据将变成纯位图格式因此与CaptureBmpImage的最终功能相同)
    function CaptureJpgImage(compressRate: Integer): IPictureDisp; safecall;
    //取得实际的视频分辨率大小
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

  //创建视频预览窗口
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


  //读取采集参数
  TCaptureParameterConfig.InitCaptureParameter(_CaptureParameter);

  stabStates.Visible := _CaptureParameter.IsShowState;

  _frmCaptureParameterCfg := nil;
end;

procedure TDSCapture.ReInitDShow;
begin
  //停止预览
  StopPreview();
  
  if Assigned(_frmCaptureParameterCfg) then FreeAndNil(_frmCaptureParameterCfg);
end;

function TDSCapture.ShowCaptureParameterCfgDialog(
  parentHandle: Integer): WideString;
begin
  try
    Result := '';

    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
      Exit;
    end;

    if not _FilterGraphic.Active then begin
      Result := 'FilterGraph尚未初始化，不能对其进行设置。';
      Exit;
    end;

    //2014-07-22 modify by tjh
    if not Assigned(_frmCaptureParameterCfg) then begin
      //需要每次重新创建参数配置窗口，因为在vb的zl9pacscapture下，如果不重新创建，
      //则第二次进入视频采集设置时，窗口将没有任何响应
      FreeAndNil(_frmCaptureParameterCfg);
      _frmCaptureParameterCfg := nil;
    end;

    //创建视频配置窗口
    _frmCaptureParameterCfg := TfrmCapParameterCfg.Create(Application);
    _frmCaptureParameterCfg.CapGraphBuilder2 := _GraphManager;
    _frmCaptureParameterCfg.CapSourceFilter := _CapSourceFilter;

    try
      _frmCaptureParameterCfg.InitParameterCfg(_CaptureParameterCfgFileName, _CaptureParameter);
    except
      on e: Exception do begin
        Application.MessageBox(PChar('初始化采集参数时产生异常，错误信息：' + e.Message), '提示', MB_OK + MB_ICONINFORMATION);
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
      //释放配置窗口对象
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

  //取得首个设备名称
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

  //取得首个设备的最大分辨率
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
  {Filter 连接图如下：

                                           /Capture Pin  --->连接录像相关Filter
  CaptureSource>>Signal Pin ---> Smart Tee                             /Capture Pin --->连接图像采集相关Filter
                                           \Preview Pin  --->Smart Tee1
                                                                       \Preview Pin --->连接视频预览相关Filter

  }
  try
    errMsg := '';

    TDebug.OutputDebug('CAP>>>Preview Step 1');
    
    _FilterGraphic.GraphEdit := _CaptureParameter.DebugFilter;

    if Trim(_CaptureParameter.CaptureDeviceName) = '' then begin

      TDebug.OutputDebug('CAP>>>Preview Step 1.1');
      //取得第一个设备名称
      _CaptureParameter.CaptureDeviceName := GetFirstDeviceName();   
      if Trim(_CaptureParameter.CaptureDeviceName) = '' then begin
        errMsg := '没有找到相关采集设备，进检查相关硬件和设置是否正确。';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 1.2');
      //取得第一个设备的最大分辨率...
      hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, _CaptureParameter.CaptureDeviceName, _CapSourceFilter);
      if Failed(hr) then begin
        errMsg := '创建CapSourceFilter视频源接口失败。 [设备名称:' + _CaptureParameter.CaptureDeviceName + ']  [错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 1.3');
      _CaptureParameter.videoSize := GetMaxVideoSize(_CapSourceFilter);
    end;
                        
    TDebug.OutputDebug('CAP>>>Preview Step 2');

    //重置采集设置
    if _FilterGraphic.Active then begin
      TDebug.OutputDebug('CAP>>>Preview Step 2.1');
      _IsPreviewState := False;

      _FilterGraphic.Stop;
      _FilterGraphic.ClearGraph; // 该过程会自动断开与filter的连接

      _FilterGraphic.Active := False;

      _VideoWindow.FilterGraph := nil;
      _ImgCaptureFilter.FilterGraph := nil;

      TDebug.OutputDebug('CAP>>>Preview Step 2.2');
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 3');

    _FilterGraphic.Active := True;


    //创建GraphManager接口对象,该对象管理_GraphBuilder中的所有FILTER
    _GraphManager := nil;
    hr := CoCreateInstance(CLSID_CaptureGraphBuilder2, nil, CLSCTX_INPROC_SERVER, IID_ICaptureGraphBuilder2, _GraphManager);
    if Failed(hr) then begin
      errMsg := '创建GraphManager接口管理对象失败。 [错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 4');

    //初始化IGraphBuilder接口对象_GraphBuilder
    hr := _GraphManager.SetFiltergraph(_FilterGraphic as IGraphBuilder);
    if Failed(hr) then begin
      errMsg := '初始化FilterGraphic对象失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 5');

    //创建视频源对象CapSourceFilter
    //if not Assigned(_CapSourceFilter) then begin
      _CapSourceFilter := nil;
      hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, _CaptureParameter.CaptureDeviceName, _CapSourceFilter);
      if Failed(hr) then begin
        errMsg := '创建CapSourceFilter视频源接口失败，请检查视频源设置。[设备名称:' + _CaptureParameter.CaptureDeviceName + '] [错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;
    //end;

    TDebug.OutputDebug('CAP>>>Preview Step 6');

    //添加_CapSourceFilter
    hr := (_FilterGraphic as IGraphBuilder).AddFilter(_CapSourceFilter, PWideChar(FILTER_NAME_CAPSOURCE_FILTER));
    if Failed(hr) then begin
      errMsg := '添加CapSourceFilter视频源接口到FilterGraphic中失败，请检查视频源设置。[设备名称:' + _CaptureParameter.CaptureDeviceName + '] [错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 7');

    //选择采集端子(已在微软视 v600 的采集卡上测试通过)
    //sdk3000在使用该段代码时，不能正常加载视频  2011-3-17
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
      //对于VFW的设备，则不进行设置

      //当视频加载的时候，视频驱动会根据之前的配置值进行加载

      //配置视频制式 （该配置需要重启预览并装载才生效）
      ConfigVideoAnalog(_CapSourceFilter, _CaptureParameter);

      TDebug.OutputDebug('CAP>>>Preview Step 8.2');
      //配置视频格式（该配置需要重启预览并装载才生效）
      ConfigVideoFormat(_CapSourceFilter, _CaptureParameter);

      TDebug.OutputDebug('CAP>>>Preview Step 8.3');
      //配置视频质量
      ConfigVideoQuality(_CapSourceFilter, _CaptureParameter);

      TDebug.OutputDebug('CAP>>>Preview Step 8.4');
    end;
    //}

    TDebug.OutputDebug('CAP>>>Preview Step 9');

    //创建SampleGrabber图像采集对象------(使用该接口的时候，不能对图像进行捕获, 因而使用DSPACK提供的TSampleGrabber对象)
    {hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_SampleGrabber, _SampleGrabber, FILTER_NAME_SAMPLE_GRABBER);
    if Failed(hr) then begin
      errMsg := '创建SampleGrabber图像采集接口失败。';
      Exit;
    end;}

    //创建SmartTee分流接口,并加入到FilterGraphic中------
    //(注：当使用RenderStream建立FILTER的连接时,如果连接的PREVIEW端口不存在，会自动创建SmartTee，
    //但有些采集设备虽然有preview的端口，却不能输出数据，不能执行filter之间的连接，当使用了SMARTTEE filter后不能作为renderstream方法的源参数，
    //估计是因为SmartTee的输出Pin并非采用PinCategory.Capture或PinCategory.Preview模式)
    _SmartTee := nil;
    _SmartTee1 := nil;
    _ColorSpace := nil;

    TDebug.OutputDebug('CAP>>>Preview Step 10');

    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_SmartTee, _SmartTee, FILTER_NAME_SMART_TEE);
    if Failed(hr) then begin
      errMsg := '创建SmartTeeFilter(0)分流接口,并加入到FilterGraphic中时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 11');
    //此段代码针对天朗采集卡PX1000E，这种采集卡需要连接两个SmartTee1的preview才能是显示的视频不卡
    //PX1000E Begin
    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_SmartTee, _SmartTee1, FILTER_NAME_SMART_TEE1);
    if Failed(hr) then begin
      errMsg := '创建SmartTeeFilter(1)分流接口,并加入到FilterGraphic中时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 12');
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _CapSourceFilter, _SmartTee, False, 0);
    if Failed(hr) then begin
      errMsg := '创建CapSourceFilter与SmartTeeFilter(0)之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 13');
    //PX1000E End

    //连接CapSourceFilter和SmartTee, 如果使用RenderStream方法，将会再添加一个smartTee接口
    //该方法将连接采集pin
    //hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _CapSourceFilter, _SmartTee, False);
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _SmartTee1, True, 1);
    if Failed(hr) then begin
      errMsg := '创建SmartTeeFilter(0)与SmartTeeFilter(1)之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 14');

    {hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _SmartTee1, True, 1);
    if Failed(hr) then begin
      errMsg := '创建ColorSpaceConvert与SmartTeeFilter之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;
    //}

    //ok c20a 2010-07-02   仿照amcap的RenderStream连接方法创建filter之间的连接，对于有些采集卡
    //smarttee Capture Pin在输出端子中序号为0，Preview Pin在输出端子中序号为1
    //如ok c20a 和 micro view v500 的preview pin就不能输出预览数据
    {hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Interleaved, _CapSourceFilter, nil, _SmartTee);
    if hr <> NOERROR then begin
      //RenderStream智能连接------(当使用RenderStream建立FILTER的连接时，会自动创建并连接SmartTee)
      hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Video, _CapSourceFilter, nil, _SmartTee);
    end;

    if hr <> NOERROR then begin
      errMsg := '建立CapSourceFilter与SmartTee之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;
    //}
            


    //判断是否动态采集动态视频
    if isCaptureVideo then begin

      TDebug.OutputDebug('CAP>>>Preview Step 14.1');
      //创建视频编码器接口
      _EncoderFilter := nil;
      hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoCompressorCategory, _CaptureParameter.EncoderName, _EncoderFilter);
      if Failed(hr) then begin
        errMsg := '创建EncoderFilter视频编码器失败，请检查编码器设置。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.2');
      //添加EncoderFilter视频编码器接口
      hr := (_FilterGraphic as IGraphBuilder).AddFilter(_EncoderFilter, PWideChar(FILTER_NAME_ENCODER));
      if Failed(hr) then begin
        errMsg := '添加EncoderFilter视频编码器接口到FilterGraphic中失败，请检查编码器设置。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.3');
      //创建MULTIPLEXER多路复用接口
      _AviMultiplexer := nil;
      hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_AviDest, _AviMultiplexer, FILTER_NAME_AVI_MULTIPLEXER);
      if Failed(hr) then begin
        errMsg := '创建AviMultiplexerFilter多路复用接口,并加入到FilterGraphic中时失败。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.4');
      //创建文件写入接口
      _AviWriter := nil;
      hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_FileWriter, _AviWriter, FILTER_NAME_AVI_WRITER);
      if Failed(hr) then begin
        errMsg := '创建AviWriterFilter文件写入接口,并加入到FilterGraphic中时失败。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.5');
      //查询视频文件设置接口
      fs := nil;
      hr := _AviWriter.QueryInterface(IID_IFileSinkFilter, fs);
      if Failed(hr)then begin
        errMsg := '查询AviWriterFilter的IFileSinkFilter接口时失败。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.6');
      //设置视频文件路径
      hr := fs.SetFileName(PWideChar(_TempCaptureVideoFile), nil);
      if FAILED(hr) then begin
        errMsg := '设置视频文件路径时失败。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.7');
      //创建SmartTee与EncoderFilter之间的连接------（使用RenderStream方式进行连接_SmartTee不能作为source filter参数）
      //使用第一个smarttee对象的capture端口进行连接，因为capture端口没有取消时间戳
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _EncoderFilter, false, 0);
      if Failed(hr) then begin
        errMsg := '创建SmartTeeFilter(1)与EncoderFilter之间的连时失败。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.8');
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _EncoderFilter, _AviMultiplexer, false, 0);
      if Failed(hr) then begin
        errMsg := '创建EncoderFilter与AviMultiplexerFilter之间的连时失败，请检查编码器设置。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.9');
      //使用RenderStream建立FILTER的连接时,会根据需要自动插入并连接SmartTee进行分流处理,但如果有PREVIEW端口时，SmartTee将不会自动插入
      //如果是用ConnectFilters的形式进行连接，则需要分别对这几个FILTER进行连接
      {hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_VIDEO, _CapSourceFilter, _EncoderFilter, _AviMultiplexer);
      if Failed(hr) then begin
        errMsg := '创建SmartTee与AviMultiplexer之间的连,并加入EncoderFilter视频编码接口时失败,请尝试重新设置视频编码器。';
        Exit;
      end;}

      //创建AviMultiplexer与AviWriter之间的连接------（需要对该连接进行处理，因为AviMultiplexer和AviWriter是由外部创建的FILTER）
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _AviMultiplexer, _AviWriter, false, 0);
      if Failed(hr) then begin
        errMsg := '创建AviMultiplexerFilter与AviWriterFilter之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>Preview Step 14.10');
    end;  //video capture filter link end...


    TDebug.OutputDebug('CAP>>>Preview Step 15');

    //根据不同的抓图方式设置显示模式
    if _CaptureParameter.SnatchWay = swVMR then begin
      TDebug.OutputDebug('CAP>>>Preview Step 15.1.1');
      _VideoWindow.Mode := vmVMR;
      _VideoWindow.VMROptions.Mode := vmrWindowless;

      _VideoWindow.FilterGraph := _FilterGraphic;

      {当直接使用vmr9接口获取图像时，需要较长的时间，因此连接_ImgCaptureFilter接口进行图像采集,
      //如果是视频回放某些格式好像不支持使用ISampleGrabber接口对象进行采集 }
      {_ImgCaptureFilter.FilterGraph := nil;

      //创建SmartTee与VideoWindow之间的连接(使用vmr模式，则不需要_ImgCaptureFilter进行图像的采集)
      hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee, _VideoWindow as IBaseFilter, true);
      if Failed(hr) then begin
        errMsg := '创建SmartTee与VideoWindow之间的连时失败。[错误代码:' + IntToStr(hr) + ']';
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


    //创建MjpegDescompress压缩接口
    {_MjpegDescompress := nil;
    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_MjpegDec, _MjpegDescompress, FILTER_NAME_MJPEGDECOMPRESS);
    if Failed(hr) then begin
      errMsg := '创建MJPEG压缩接口,并加入到FilterGraphic中时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end; }



    //创建SmartTee与ImgCaptureFilter之间的连接  (当使用窗口模式时，需要连接_ImgCaptureFilter)
    //将下句代码，修改为使用采集信号脚，直接连接图像捕获_ImgCaptureFilter的filter，
    //hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee1, _ImgCaptureFilter  as IBaseFilter, true, 1);

    //针对部分采集卡，在安装一些编码器后，smarttee1和imacapturefilter之间会自动插入安装后的编码器，造成采集图像时偏色，
    //因此需要在此之间手动插入color space converter进行转换

    TDebug.OutputDebug('CAP>>>Preview Step 17');
    //创建ColorSpaceConverter Filter避免预览所见和采集的图像颜色有所偏差
    hr := TDS9Ex.AddFilterToGraphBuilder(_FilterGraphic as IGraphBuilder, CLSID_Colour, _ColorSpace, FILTER_NAME_COLOR_CONVERT);
    if Failed(hr) then begin
      errMsg := '创建ColorSpaceConvert颜色空间转换接口,并加入到FilterGraphic中时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 18');

    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee1, _ColorSpace, false, 0);
    if Failed(hr) then begin
      errMsg := '创建SmartTeeFilter(1)与ColorSpaceConverter之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 19');

    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _ColorSpace, _ImgCaptureFilter  as IBaseFilter, false, 0);
    if Failed(hr) then begin
      errMsg := '创建ColorSpaceConverter与ImgCaptureFilter之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end;

    TDebug.OutputDebug('CAP>>>Preview Step 20');
    {//不能成功连接到MjpegDescompressFilter,GraphiEdit会自动添加AVICompressorFilter
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _MjpegDescompress, _ImgCaptureFilter as IBaseFilter, false);
    if Failed(hr) then begin
      errMsg := '创建ImgCaptureFilter与MjpegDescompressFilter之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end; //}

    //创建ImgCaptureFilter与VideoWindow之间的连接
    //连接videowindow进行显示时，直接使用预览脚输出信号显示。
    //hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _ImgCaptureFilter as IBaseFilter, _VideoWindow as IBaseFilter, false, 0);
    hr := TDS9Ex.ConnectFilters(_FilterGraphic as IGraphBuilder, _SmartTee1, _VideoWindow as IBaseFilter, true, 1);
    if Failed(hr) then begin
      errMsg := '创建ImgCaptureFilter与VideoWindow之间的连接时失败。[错误代码:' + IntToStr(hr) + ']';
      Exit;
    end; //}

    TDebug.OutputDebug('CAP>>>Preview Step 21');
    //RenderStream使用智能连接(对于有些采集设备来说虽然具备PREVIEW端口，但却不能使用该端口输出数据)
    {hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_PREVIEW, @MEDIATYPE_VIDEO, _CapSourceFilter, _ImgCaptureFilter as IBaseFilter, _VideoWindow as IBaseFilter);
    if Failed(hr) then begin
      hr := (_FilterGraphic as ICaptureGraphBuilder2).RenderStream(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_VIDEO, _CapSourceFilter, _ImgCaptureFilter as IBaseFilter, _VideoWindow as IBaseFilter);
      if Failed(hr) then begin
        errMsg := '创建CapSourceFilter与ImgCaptureFilter之间的连接时失败，不能预览视频图像，请检查设备的输出端口类型。';
        Exit;
      end;
    end;}

    _FilterGraphic.Play;

    TDebug.OutputDebug('CAP>>>Preview Step 22');
    imgLogo.Visible := False;

    //调整窗口位置
    AdjustWindowSize();

    TDebug.OutputDebug('CAP>>>Preview Step 23');
    try
      if Assigned(_frmCaptureParameterCfg) then begin
        //需要对参数设置窗口中的采集相关filter进行更新
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
    //如果采集设备为空，则不进行设置
    if Trim(capParameter.CaptureDeviceName) = '' then Exit;

    //判断是否进入预览模式
    if not _IsPreviewState then begin
      TCaptureParameterConfig.CopyParameter(capParameter, _CaptureParameter);

      //在设置改编的时候，如果没有预览，则执行开始预览
      if Trim(_CaptureParameter.CaptureDeviceName) <> '' then
        StartPreview();

      Exit;
    end;

    if needCaptureSample then begin
      //采集样品图像用于裁剪设置
      tmpBitMap := CaptureImageToBmpObj;
      try
        tmpBitMap.SaveToFile(TfrmCapParameterCfg.GetCaptureSampleFile);
      finally
        FreeAndNil(tmpBitMap);
      end;

      Exit;
    end;

    //判断参数是否修改，如果修改则更新视频显示**************************************

    //更新采集设备
    if capParameter.CaptureDeviceName <> _CaptureParameter.CaptureDeviceName then begin
      _CaptureParameter.CaptureDeviceName := capParameter.CaptureDeviceName;
      Preview(False, previewResult);
    end;

    //刷新视频质量
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

      //如果是VFW设备，则不进行调整
      if not TDS9Ex.IsVfwDevice(capParameter.CaptureDeviceName) then
        ConfigVideoQuality(_CapSourceFilter, _CaptureParameter);
    end;

    //刷新视频格式
    if (capParameter.VideoSize <> _CaptureParameter.VideoSize) then begin
      _CaptureParameter.VideoSize := capParameter.VideoSize;

      Preview(False, previewResult);
      //if FEvents <> nil then FEvents.OnResize;//在VB中，运行的时候，该事件才可以被正常调用执行
      VideoSizeEvent(_VideoWindow.Width, _VideoWindow.Height, Self.Width, Self.Height);
    end;

    //刷新颜色深度
    if (capParameter.ColorDepth <> _CaptureParameter.ColorDepth) then begin
      _CaptureParameter.ColorDepth := capParameter.ColorDepth;
      Preview(false, previewResult);
    end;

    //刷新视频制式
    if capParameter.VideoAnalog <> _CaptureParameter.VideoAnalog then begin
      _CaptureParameter.VideoAnalog := capParameter.VideoAnalog;
      Preview(false, previewResult);
    end;

    //刷新显示模式
    if capParameter.VideoShowModel <> _CaptureParameter.VideoShowModel then begin
      _CaptureParameter.VideoShowModel := capParameter.VideoShowModel;
      AdjustWindowSize();
    end;

    //刷新图像抓取模式
    if capParameter.SnatchWay <> _CaptureParameter.SnatchWay then begin
      _CaptureParameter.SnatchWay := capParameter.SnatchWay;

      if _IsPreviewState and not _IsCaptureVideo then begin
        //重新开始预览
        Preview(False, sErrMsg);
      end;
    end;

    //刷新输入端口
    if capParameter.InputCrossbar <> _CaptureParameter.InputCrossbar then begin
      _CaptureParameter.InputCrossbar := capParameter.InputCrossbar;
      if _IsPreviewState and not _IsCaptureVideo then begin
        //重新开始预览
        Preview(False, sErrMsg);      
      end;
    end;

    //刷新输出端口
    if capParameter.OutputCrossbar <> _CaptureParameter.OutputCrossbar then begin
      _CaptureParameter.OutputCrossbar := capParameter.OutputCrossbar;
      if _IsPreviewState and not _IsCaptureVideo then begin
        //重新开始预览
        Preview(False, sErrMsg);      
      end;
    end;


    //视频状态显示设置
    if capParameter.IsShowState <> _CaptureParameter.IsShowState then begin
      _CaptureParameter.IsShowState := capParameter.IsShowState;

      stabStates.Visible := capParameter.IsShowState;
      AdjustWindowSize();
    end;

    //复制其他不需要刷新当前视频显示的参数
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
  curBaseFilter := filter; //filter可以修改为TFilter类
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

  //取得质量配置相关信息
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
      //取得视频质量设置的范围
      curHr := curAmVideoProcAmp.GetRange(PropertyTag, iMinValue, iMaxValue, iStep, iDefault, iFlags);
      if not Succeeded(curHr) then begin
        isSucceed := False;
        Result := 0;
        TDebug.OutputDebug('CAP>>>GetDefaultQualityInf 1.1 Return False');
        Exit;
      end;

      TDebug.OutputDebug('CAP>>>GetDefaultQualityInf 2');
      //取得当前值
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

      //tagCameraControlFlags中定义的值与微软声明相反，参考定义注释部分
      amCameraControl.Set_(CameraControl_Exposure, captureParameter.ExposureValue, TCameraControlFlags(captureParameter.ExposureWay));
      TDebug.OutputDebug('CAP>>>ConfigVideoQuality 1.5 auto Exposure');
    end;
  end;

  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 2');
  hr := curBaseFilter.QueryInterface(IID_IAMVideoProcAmp, amVideoProcAmp);
  if not Succeeded(hr) then Exit;

  //说明：经测试在directshow中VideoProcAmp_Flags_Auto表示手动管理，  VideoProcAmp_Flags_Manual表示自动，产生此问题是因为值定义错误

  isOk := False;
  
  TDebug.OutputDebug('CAP>>>ConfigVideoQuality 3');
  //亮度
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
  //对比度
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
  //色调
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
  //饱和度
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
  //伽马
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
  //白平衡
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
      Result := '视频采集尚未进入预览模式。';
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := '没有采集到视频图像。';
        Exit;
      end;

      //转换为灰度图
      //采集的图像一般都在1024*768以内，所以即便是不经过裁剪进行灰度转换，
      //在效率上也没有多大的影响
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //图像裁剪操作
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //取得裁剪范围
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
        //直接保存采集图像
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
    Result := '视频采集尚未进入预览模式。';
    Exit;
  end;

  if _IsCaptureVideo then begin
    Result := '正在进行视频采集，不能执行该操作。';
    Exit;
  end;

  try
    //设置视频文件保存位置
    _TempCaptureVideoFile := fileName;
    if Trim(fileName) = '' then begin
      CreateGUID(fileGuid);
      curFile := GUIDToString(fileGuid) + '.avi';
      curFile := StringReplace(curFile, '-', '', [rfReplaceAll, rfIgnoreCase]);

      _TempCaptureVideoFile := ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR + curFile;

      //如果目录不存在，则创建该目录
      if not DirectoryExists(ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR) then begin
        ForceDirectories(ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR);
      end;
    end;

    //开始动态视频采集
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
    if not _IsCaptureVideo then Exit; //没有开始视频采集，则不执行该操作。 

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
        //采集时间限制(不能超过指定时间),如果未进行设置，则不允许超过8小时（3600 * 8）
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

  //需要调整视频显示位置   
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
      stabStates.Panels.Items[3].Text := '预览模式';
    end else begin
      stabStates.Panels.Items[3].Text := '录像模式';
    end;

    if not _IsPreviewState then begin
      stabStates.Panels.Items[3].Text := '闲置模式';
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
  //根据显示模式类型，设置相关位置及大小
  case _CaptureParameter.VideoShowModel of
    smNormal: begin //---------------------------------------------------------
      _VideoWindow.Align := alNone;
      curVideoSizeInf := TCaptureParameterConfig.ConvertVideoSizeInf(_CaptureParameter.VideoSize);      

      _VideoWindow.Width := curVideoSizeInf.Width;
      _VideoWindow.Height := curVideoSizeInf.Height;

      _VideoWindow.Left := (Self.Width - _VideoWindow.Width) div 2 - 1;

      //判断是否需要减去状态栏高度
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


      //判断是否需要减去状态栏高度
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height;

      //取得缩放比率
      if (curVideoSizeInf.Height) / curVideoSizeInf.Width > (Self.Height - stateBarHeight) / (Self.Width) then begin
        zoomRate := (Self.Height - stateBarHeight) / curVideoSizeInf.Height;
      end else begin
        zoomRate := Self.Width / curVideoSizeInf.Width;
      end;

      //如果大小相等，则不进行缩放
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
      //判断当前窗口大小是否适应采集窗口大小
      curVideoSizeInf := TCaptureParameterConfig.ConvertVideoSizeInf(_CaptureParameter.VideoSize);

      //判断是否需要减去状态栏高度
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
    Result := '正在采集视频，不能停止预览。';
    Exit;
  end;

  try
    //停止视频预览（必须按照这样的顺序执行）
    if _FilterGraphic.Active then begin
                        
      _FilterGraphic.Stop;

      _FilterGraphic.ClearGraph; // 该过程会自动断开与filter的连接

      _FilterGraphic.Active := False;
    end;

    _GraphManager := nil;

    _IsPreviewState := False;
    imgLogo.Visible := True;

    //该方法会触发videowindow 的paint事件，因此需要放在_IsPreviewState := False语句之后。
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
  //退出全屏
  if _IsClickQuitFullScreen {and _VideoWindow.FullScreen} then begin
    QuitFullScreen();
  end;
    
  if Assigned(Self.OnClick) then Self.OnClick(Sender);
end;

procedure TDSCapture._VideoWindowDblClick(Sender: TObject);
begin
  //退出全屏
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
      //采集单帧图像     取得指定图像后再转换为指定位数的图像，dshow采集时的图像位数设置的为32位
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

        //为避免发音时，占用过程时间，因此使用postmessage方法触发声音提示
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
        //为避免发音时，占用过程时间，因此使用postmessage方法触发声音提示
        if _CaptureParameter.IsSoundHint then PostMessage(Self.Handle, WM_BEEPMSG, 0, 0);

        Exit;
      end else begin
        //当使用SampleGrabber不能采集到图像时，则直接使用VideoWindow的VMRGetBitmap方法采集图像
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

          //为避免发音时，占用过程时间，因此使用postmessage方法触发声音提示
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
      Result := '正在进行视频采集，不能执行该操作。';
      Exit;
    end;

    if Trim(_CaptureParameter.CaptureDeviceName) = '' then begin
      Result := '尚未设置采集设备名称。';
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

//显示采集pin配置
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
      Result := '尚未进入预览模式，不能执行该操作。';
      Exit;
    end;  

    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
      Exit;
    end;
      
    //该设置需要先执行停止操作
    _FilterGraphic.Stop();
    //pinlist := TPinList.Create(_CapSourceFilter);
    try
      //查找CAPTURE PIN
      //(_FilterGraphic as ICaptureGraphBuilder2).FindPin(_SourceFilter as IbaseFilter, PINDIR_OUTPUT, @PIN_CATEGORY_PREVIEW, @MEDIATYPE_Video, false, 0, pPinOut);

      //网上说pinlist.Items[0]表示预览的分辨率，pinlist.Items[1]表示捕获的分辨率
      {for i := 0 to pinlist.Count - 1 do begin
        if (pinlist.PinInfo[i].dir = PINDIR_OUTPUT) and (pinlist.Connected[i]) then begin
          TDS9Ex.ShowPinPropertyPage('视频端口', parentHandle, pinlist.Items[i]);
          exit;
        end;
      end;}

      //使用amcap的实现方式显示采集端口属性
      curVideoSize := TCaptureParameterConfig.ConvertVideoSizeInf(_captureParameter.VideoSize);

      TDS9Ex.ShowPinPropertyPage1('视频端口', parentHandle, _GraphManager, _CapSourceFilter, curVideoSize);
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
      Result := '尚未进入预览模式，不能执行该操作。';
      Exit;
    end;

    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
      Exit;
    end;
      
    //ShowVfwConfigDialog(VfwCaptureDialog_Display, parentHandle);
    TDS9Ex.ShowFilterPropertyPage('视频显示设置', parentHandle, _VideoWindow as IBaseFilter);

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
      Result := '尚未进入预览模式，不能执行该操作。';
      Exit;
    end;  
    
    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
      Exit;
    end;

    //ShowVfwConfigDialog(VfwCaptureDialog_Format, parentHandle);

    _FilterGraphic.Stop();
    try
      TDS9Ex.ShowFilterPropertyPage('视频格式', parentHandle, _CapSourceFilter as IBaseFilter, ppVFWCapSource);
      //if Succeeded(hr) then begin
      //  TDS9Ex.ShowFilterPropertyPage('视频格式', parentHandle, _CapSourceFilter as IBaseFilter, ppDefault);
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
      Result := '尚未进入预览模式，不能执行该操作。';
      Exit;
    end;

    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
      Exit;
    end;

    _FilterGraphic.Stop();

    try
      //ShowVfwConfigDialog(VfwCaptureDialog_Source, parentHandle);
      hr := TDS9Ex.ShowFilterPropertyPage('视频源', parentHandle, _CapSourceFilter as IBaseFilter, ppVFWCapSource);
      if Succeeded(hr) then begin
        TDS9Ex.ShowFilterPropertyPage('视频源', parentHandle, _CapSourceFilter as IBaseFilter, ppDefault);
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

  //该设置需要先执行停止操作
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
  raise Exception.Create('暂未实现该功能。');
end;

procedure TDSCapture.vfwConfigCallEvent(const operVfwConfigType: TVfwConfigType;
  const parentHandle: Integer; out errMsg: WideString);
begin
  case operVfwConfigType of
    vctVideoSourceProperty: begin //显示VFW源设置
      errMsg := ShowVideoCaptureFilterCfg(parentHandle);
    end;
    vctVideoCapturePinProperty: begin  //显示采集端口属性设置
      errMsg := ShowVideoCapturePinCfg(parentHandle);
    end;
    vctVfwVideoFormat: begin   //显示视频格式设置
      errMsg := ShowVfwVideoFormatCfg(parentHandle);
    end;
    vctVfwVideoDisplay: begin  //显示视频显示设置
      errMsg := ShowVfwVideoDisplayCfg(parentHandle);
    end;
    vctVideoCrossbar: begin    //显示video Crossbar设置
      errMsg := ShowVideoCrossbarCfg(parentHandle);
    end;
    vctVfwCompressDialog: begin  //显示视频压缩设置
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
      Result := '视频采集尚未进入预览模式。';
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := '没有采集到视频图像。';
        Exit;
      end;

      //转换为灰度图
      //采集的图像一般都在1024*768以内，所以即便是不经过裁剪进行灰度转换，
      //在效率上也没有多大的影响
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //图像裁剪操作
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //取得裁剪范围
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
        //直接保存采集图像
        jpg := TGraphicProcess.BmpConvertToJpg(bitMap, compressRate);
      end;

      if not Assigned(jpg) then begin
        Result := '图像转换失败。';
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
  //AdjustWindowSize(); //当显示窗口大小发生改变时，自动改变视频输出的显示位置(可由外部调用RefreshWindow更新)

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
  //退出全屏模式
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

  //当按下屏幕锁定或者任务键时，在恢复的时候刷新视频显示
  if not _IsPreviewState then Exit;

  if _CaptureParameter.SnatchWay = swVMR then begin
    //Vmr方式的刷新
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
    //device方式的刷新
    hr := (_VideoWindow as IBaseFilter).QueryInterface(IID_IVideoWindow, vw);
    if Failed(hr) then Exit;

    try
      //刷新视频
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
    //重新开始预览
    Preview(False, sErrMsg);
  end;
end;


function TDSCapture.UpdateVideoQuailty: WideString;
begin
  if not TDS9Ex.IsVfwDevice(_CaptureParameter.CaptureDeviceName)
    and _IsPreviewState then begin
    //配置视频质量对于VFW的设备，则不进行设置
    ConfigVideoQuality(_CapSourceFilter, _CaptureParameter);
  end;
end;

function TDSCapture.SaveParameterToFile: WideString;
begin
  try
    Result := '';

    //写入采集参数
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
    Result := '正在进行视频采集，不能执行该操作。';
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
      Result := '正在进行视频采集，不能执行该操作。';
      Exit;
    end;

    if Trim(_CaptureParameter.encoderName) = '' then begin
      Result := '尚未设置视频编码器名称。';
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
        Result := '没有采集到视频图像。';
        Exit;
      end;

      //转换为灰度图
      //采集的图像一般都在1024*768以内，所以即便是不经过裁剪进行灰度转换，
      //在效率上也没有多大的影响
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //图像裁剪操作
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //取得裁剪范围
          cutArea.Left := Round(_CaptureParameter.LeftRate * bitMap.Width);
          cutArea.Right := cutArea.Left + Round(_CaptureParameter.WidthRate * bitMap.Width);

          cutArea.Top := Round(_CaptureParameter.TopRate * bitMap.Height);
          cutArea.Bottom := cutArea.Top + Round(_CaptureParameter.HeightRate * bitMap.Height);

          TGraphicProcess.CutImg(cutArea, bitMap, cutBmp);



          cutBmp.SaveToClipboardFormat(curFormat, data, palette);

          {curFormat := format;
          if curFormat <= 0 then begin
            curFormat := RegisterClipboardFormat('ZLDSVIDEOPROCESS10161');

            H:= GlobalAlloc(GMEM_DDESHARE, GlobalSize(data)); //分配一块内存
            //P:= GlobalLock(H);//取得指向内存块的指针

            //move(Pchar('')^, p^, GlobalSize(data));

            //GlobalUnlock(H);

          end;}

          Clipboard.SetAsHandle(curFormat, data);
          //format := curFormat;
        finally
          FreeAndNil(cutBmp);
        end;
      end else begin
        //直接将图像复制到剪贴板
        bitMap.SaveToClipboardFormat(curFormat, data, palette);

        {curFormat := format;
        if curFormat <= 0 then begin
          curFormat := RegisterClipboardFormat('ZLDSVIDEOPROCESS10161');

          H:= GlobalAlloc(GMEM_DDESHARE, GlobalSize(data)); //分配一块内存
          //P:= GlobalLock(H);//取得指向内存块的指针

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
      Result := '尚未进入预览模式，不能执行该操作。';
      Exit;
    end;

    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
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
      Result := '尚未进入预览模式，不能执行该操作。';
      Exit;
    end;  

    if _IsCaptureVideo then begin
      Result := '正在进行视频采集，不能对其进行设置。';
      Exit;
    end;

    //该设置需要先执行停止操作  在amcap中没有停止filter graph 但是在具体测试的时候，
    //如果不停止，有可能第一次将弹出错误提示，或者没有任何反映
    _FilterGraphic.Stop();
    try

      TDS9Ex.ShowVideoCrossbarPropertyPage('视频端口', parentHandle, _GraphManager, _CapSourceFilter);
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

      //转换为灰度图
      //采集的图像一般都在1024*768以内，所以即便是不经过裁剪进行灰度转换，
      //在效率上也没有多大的影响
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //图像裁剪操作
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //取得裁剪范围
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

      //直接保存采集图像
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

      //转换为灰度图
      //采集的图像一般都在1024*768以内，所以即便是不经过裁剪进行灰度转换，
      //在效率上也没有多大的影响
      if _CaptureParameter.IsConvertGrayImg then begin
        TGraphicProcess.ConvertBitmapToGrayscale(bitMap);
      end;


      //图像裁剪操作
      if _CaptureParameter.IsApplyImageCut then begin
        cutBmp := TBitmap.Create;
        try
          //取得裁剪范围
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

      //保存jpg图像
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
    //查找已经连接的PIN
    hr := TDS9Ex.FindConnectedPin(_CapSourceFilter, PINDIR_OUTPUT, pin);
    if FAILED(hr) then Exit;
    if not Assigned(pin) then Exit;

    try                             
      hr := pin.QueryInterface(IID_IAMStreamConfig, amStreamConfig);
      if FAILED(hr) then Exit;

      try
        hr := amStreamConfig.GetFormat(pmt);   //取得默认视频格式
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
