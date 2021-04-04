unit DSPlayImpl1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ActiveX, AxCtrls, ZLDSVideoProcess_TLB, StdVcl, ComCtrls, DSPack,
  ExtCtrls, VideoProcessDefine;


const
  WM_BEEPMSG = wm_user + $1088;


type
  TDSPlay = class(TActiveForm, IDSPlay)
    VideoWindow: TVideoWindow;
    stabStates: TStatusBar;
    FilterGraph: TFilterGraph;
    timerState: TTimer;
    ImgCapture: TSampleGrabber;
    imgLogo: TImage;
    imgAnimate: TImage;
    procedure VideoWindowPaint(Sender: TObject);
    procedure VideoWindowKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure VideoWindowClick(Sender: TObject);
    procedure VideoWindowDblClick(Sender: TObject);
    procedure VideoWindowEnter(Sender: TObject);
    procedure VideoWindowExit(Sender: TObject);
    procedure VideoWindowKeyPress(Sender: TObject; var Key: Char);
    procedure VideoWindowKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure VideoWindowMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure VideoWindowMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure VideoWindowMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure VideoWindowMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure VideoWindowMouseWheelDown(Sender: TObject;
      Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
    procedure VideoWindowMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure timerStateTimer(Sender: TObject);
  private
    { Private declarations }
    _VideoState: TVideoState;      //��ǰ��Ƶ����״̬
    _rateStep: Double;             //���ӻ��߼��ٵ����ʲ���

    _VideoInf: TVideoInf;          //������Ƶ�Ļ�����Ϣ
    _ShowModel: TShowModel;        //���Ŵ�����ʾģʽ
    
    _IsFit: Boolean;
    _IsStretch: Boolean;
    _IsAdjustWindowSize: Boolean;  //�����Ƿ��Զ���Ӧ�ֱ��ʴ�С

    _IsEscKeyQuitFullScreen: Boolean;
    _IsDblClickQuitFullScreen: Boolean;
    _IsClickQuitFullScreen: Boolean;    

    _VideoFile: WideString;         //��ǰ���ŵ���Ƶ�ļ�����

    _SnatchWay: TSnatchWay;         //ͼ��ץȡ��ʽ��������Ҳ��������Ƶ����ʾģʽ�������ֵΪswVMR��ʹ��VMR��windowlessģʽ��
    _IsSoundHint: Boolean;          //�ɼ�ͼ��ʱ���Ƿ����������ʾ
    _IsDebugFilter: Boolean;        //�Ƿ��Filter���е��� 

    FEvents: IDSPlayEvents;
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
    procedure MouseWheelEvent(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure MouseWheelDownEvent(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure MouseWheelUpEvent(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);      

    procedure InitDSPlay();
    procedure ReInitDSPlay();

    //��ȡ��Ƶ������Ϣ
    procedure InitVideoInf(const fileName: WideString);

    //������Ƶ������ʾ��ʽ
    procedure AdjustWindowSize();
    //�ɼ�BMPͼ��
    function CaptureImageToBmpObj: TBitmap;
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

    //custom property
    function Get_IsFit: WordBool; safecall;
    function Get_IsFullScreen: WordBool; safecall;
    function Get_IsStretch: WordBool; safecall;
    procedure Set_IsFit(Value: WordBool); safecall;
    procedure Set_IsFullScreen(Value: WordBool); safecall;
    procedure Set_IsStretch(Value: WordBool); safecall;
    function Get_CurFrame: SYSINT; safecall;
    function Get_CurTime: SYSINT; safecall;
    function Get_FrameLen: SYSINT; safecall;
    function Get_TimeLen: SYSINT; safecall;
    procedure Set_CurFrame(Value: SYSINT); safecall;
    procedure Set_CurTime(Value: SYSINT); safecall;
    function Get_PlayRate: Double; safecall;
    procedure Set_PlayRate(Value: Double); safecall;
    function Get_VideoState: TVideoState; safecall;
    function Get_ShowModel: TShowModel; safecall;
    procedure Set_ShowModel(Value: TShowModel); safecall;
    function Get_IsAdjustWindowSize: WordBool; safecall;
    function Get_IsShowState: WordBool; safecall;
    procedure Set_IsAdjustWindowSize(Value: WordBool); safecall;
    procedure Set_IsShowState(Value: WordBool); safecall;
    function Get_IsClickQuitFullScreen: WordBool; safecall;
    function Get_IsDblClickQuitFullScreen: WordBool; safecall;
    function Get_IsEscKeyQuitFullScreen: WordBool; safecall;
    procedure Set_IsClickQuitFullScreen(Value: WordBool); safecall;
    procedure Set_IsDblClickQuitFullScreen(Value: WordBool); safecall;
    procedure Set_IsEscKeyQuitFullScreen(Value: WordBool); safecall;
    function Get_CurHeight: Integer; safecall;
    function Get_CurWidth: Integer; safecall;
    procedure Set_CurHeight(Value: Integer); safecall;
    procedure Set_CurWidth(Value: Integer); safecall;
    function Get_SnatchWay: TSnatchWay; safecall;
    procedure Set_SnatchWay(Value: TSnatchWay); safecall;
    function Get_AppHandle: Integer; safecall;
    procedure Set_AppHandle(Value: Integer); safecall;
    function Get_Balance: Integer; safecall;
    function Get_Volume: Integer; safecall;
    procedure Set_Balance(Value: Integer); safecall;
    procedure Set_Volume(Value: Integer); safecall;
    function Get_StreamTypeName: WideString; safecall;
    function Get_IsSoundHint: WordBool; safecall;
    procedure Set_IsSoundHint(Value: WordBool); safecall;
    function Get_IsDebugFilter: WordBool; safecall;
    procedure Set_IsDebugFilter(Value: WordBool); safecall;
    function Get_VideoFile: WideString; safecall;
    procedure Set_VideoFile(const Value: WideString); safecall;      

    //��ͣ����
    function Pause: WideString; safecall;
    //������Ƶ�ļ�
    function Play(const videoFile: WideString): WideString; safecall;
    //ֹͣ����
    function Stop: WideString; safecall;
    //��������
    function Run: WideString; safecall;
    //���ٲ���
    function AddRate: WideString; safecall;
    //�ɼ�BMPͼ�񵽴���
    function CaptureBmpImgToFile(const fileName: WideString): WideString; safecall;
    //�ɼ�JPGͼ�񵽴���
    function CaptureJpgImgToFile(const fileName: WideString; compressRate: SYSINT): WideString; safecall;
    //���ٲ���
    function DecRate: WideString; safecall;
    //�ָ�������������
    function RestoreRate: WideString; safecall;
    //��ʾ��Ƶ��Ϣ
    function ShowVideoInfo(parentHandle: SYSINT): WideString; safecall;
    //�ͷ���Դ
    procedure FreeRes; safecall;
    //��һ֡
    function FirstFrame: WideString; safecall;
    //���һ֡
    function LastFrame: WideString; safecall;
    //��һ֡
    function NextFrame: WideString; safecall;
    //��һ֡
    function PriorFrame: WideString; safecall;
    //�˳�ȫ����ʾ
    function QuitFullScreen: WideString; safecall;
    //��ʾȫ��
    function ShowFullScreen(parentHandle, monitorIndex: Integer): WideString; safecall;
    //ˢ�´���
    function RefreshWindow: WideString; safecall;
    //ȡ���������
    function GetVideoProperty(propertyType: TVideoProperty; var value: WideString): WideString; safecall;
    //�ظ�����
    function RePlay: WideString; safecall;
    //�ɼ�ͼ�񵽼�����
    function CaptureImgToClipBoard: WideString; safecall;
    //��ʾ��ͨӰƬ
    procedure ShowAnimate(AnimateType: TAnimateType); safecall;
    //���ؿ�ͨӰƬ
    procedure HideAnimate; safecall;
    //�ɼ�bmpͼ�񣬷���IPictureDisp�ӿ�
    function CaptureBmpImage: IPictureDisp; safecall;

    procedure WM_BEEP(var msg: TMessage); message WM_BEEPMSG;
  public
    { Public declarations }
    procedure Initialize; override;
  end;

implementation

uses
  ComObj, ComServ, DirectShow9, Jpeg, DirectShow9Ex, GraphicProcess, FullScreenWindow,
  CaptureDebug, VideoInfWindow, Clipbrd, GifImage;

{$R *.DFM}
//{$R Soundgif.RES}

{ TDSPlay }

procedure TDSPlay.DefinePropertyPages(DefinePropertyPage: TDefinePropertyPage);
begin
  { Define property pages here.  Property pages are defined by calling
    DefinePropertyPage with the class id of the page.  For example,
      DefinePropertyPage(Class_DSPlayPage); }
end;

procedure TDSPlay.EventSinkChanged(const EventSink: IUnknown);
begin
  FEvents := EventSink as IDSPlayEvents;
  inherited EventSinkChanged(EventSink);
end;

procedure TDSPlay.Initialize;
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

  {
  VideoWindow.OnMouseMove := MouseMoveEvent;
  VideoWindow.OnMouseDown := MouseDownEvent;
  VideoWindow.OnMouseUp := MouseUpEvent;
  VideoWindow.OnEnter := EnterEvent;
  VideoWindow.OnExit := ExitEvent;
  }

  InitDSPlay();
end;

function TDSPlay.Get_Active: WordBool;
begin
  Result := Active;
end;

function TDSPlay.Get_AlignDisabled: WordBool;
begin
  Result := AlignDisabled;
end;

function TDSPlay.Get_AutoScroll: WordBool;
begin
  Result := AutoScroll;
end;

function TDSPlay.Get_AutoSize: WordBool;
begin
  Result := AutoSize;
end;

function TDSPlay.Get_AxBorderStyle: TxActiveFormBorderStyle;
begin
  Result := Ord(AxBorderStyle);
end;

function TDSPlay.Get_Caption: WideString;
begin
  Result := WideString(Caption);
end;

function TDSPlay.Get_Color: OLE_COLOR;
begin
  Result := OLE_COLOR(Color);
end;

function TDSPlay.Get_DoubleBuffered: WordBool;
begin
  Result := DoubleBuffered;
end;

function TDSPlay.Get_DropTarget: WordBool;
begin
  Result := DropTarget;
end;

function TDSPlay.Get_Enabled: WordBool;
begin
  Result := Enabled;
end;

function TDSPlay.Get_Font: IFontDisp;
begin
  GetOleFont(Font, Result);
end;

function TDSPlay.Get_HelpFile: WideString;
begin
  Result := WideString(HelpFile);
end;

function TDSPlay.Get_KeyPreview: WordBool;
begin
  Result := KeyPreview;
end;

function TDSPlay.Get_PixelsPerInch: Integer;
begin
  Result := PixelsPerInch;
end;

function TDSPlay.Get_PrintScale: TxPrintScale;
begin
  Result := Ord(PrintScale);
end;

function TDSPlay.Get_Scaled: WordBool;
begin
  Result := Scaled;
end;

function TDSPlay.Get_ScreenSnap: WordBool;
begin
  Result := ScreenSnap;
end;

function TDSPlay.Get_SnapBuffer: Integer;
begin
  Result := SnapBuffer;
end;

function TDSPlay.Get_Visible: WordBool;
begin
  Result := Visible;
end;

function TDSPlay.Get_VisibleDockClientCount: Integer;
begin
  Result := VisibleDockClientCount;
end;

procedure TDSPlay._Set_Font(var Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TDSPlay.ActivateEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnActivate;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'ActivateEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.ClickEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnClick;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'ActivateEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.CreateEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnCreate;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'CreateEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.DblClickEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnDblClick;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'DblClickEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.DeactivateEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnDeactivate;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'DeactivateEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.DestroyEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnDestroy;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'DestroyEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.KeyPressEvent(Sender: TObject; var Key: Char);
var
  TempKey: Smallint;
begin
  try
    TempKey := Smallint(Key);
    if FEvents <> nil then FEvents.OnKeyPress(TempKey);
    Key := Char(TempKey);
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'KeyPressEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.PaintEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnPaint;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'PaintEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.Set_AutoScroll(Value: WordBool);
begin
  AutoScroll := Value;
end;

procedure TDSPlay.Set_AutoSize(Value: WordBool);
begin
  AutoSize := Value;
end;

procedure TDSPlay.Set_AxBorderStyle(Value: TxActiveFormBorderStyle);
begin
  AxBorderStyle := TActiveFormBorderStyle(Value);
end;

procedure TDSPlay.Set_Caption(const Value: WideString);
begin
  Caption := TCaption(Value);
end;

procedure TDSPlay.Set_Color(Value: OLE_COLOR);
begin
  Color := TColor(Value);
  VideoWindow.Color := TColor(Value);
end;

procedure TDSPlay.Set_DoubleBuffered(Value: WordBool);
begin
  DoubleBuffered := Value;
end;

procedure TDSPlay.Set_DropTarget(Value: WordBool);
begin
  DropTarget := Value;
end;

procedure TDSPlay.Set_Enabled(Value: WordBool);
begin
  Enabled := Value;
end;

procedure TDSPlay.Set_Font(const Value: IFontDisp);
begin
  SetOleFont(Font, Value);
end;

procedure TDSPlay.Set_HelpFile(const Value: WideString);
begin
  HelpFile := String(Value);
end;

procedure TDSPlay.Set_KeyPreview(Value: WordBool);
begin
  KeyPreview := Value;
end;

procedure TDSPlay.Set_PixelsPerInch(Value: Integer);
begin
  PixelsPerInch := Value;
end;

procedure TDSPlay.Set_PrintScale(Value: TxPrintScale);
begin
  PrintScale := TPrintScale(Value);
end;

procedure TDSPlay.Set_Scaled(Value: WordBool);
begin
  Scaled := Value;
end;

procedure TDSPlay.Set_ScreenSnap(Value: WordBool);
begin
  ScreenSnap := Value;
end;

procedure TDSPlay.Set_SnapBuffer(Value: Integer);
begin
  SnapBuffer := Value;
end;

procedure TDSPlay.Set_Visible(Value: WordBool);
begin
  Visible := Value;
end;

function TDSPlay.Pause: WideString;
begin
  try
    Result := '';

    //���Ϊֹͣ״̬�������޸�Ϊ��ͣ״̬
    if _VideoState = vsStop then Exit;

    FilterGraph.Pause;

    _VideoState := vsPause;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.Play(const videoFile: WideString): WideString;
//Ŀǰ���ܶ�asf��ʽ����Ƶ���е�֡ͼ��Ĳɼ�
var
  hr: HRESULT;
  vwRect: TRect;
    
  vw: IVideoWindow;
  vmr9windLessCtrl: IVMRWindowlessControl9;

  function CheckIsSupportVmr: WordBool;
  var
   AFilter: IBaseFilter;
   CW: Word;
  begin
    try
      CW := Get8087CW;
      try
        result := (CoCreateInstance(CLSID_VideoMixingRenderer9, nil, CLSCTX_INPROC, IID_IBaseFilter ,AFilter) = S_OK);
      finally
        Set8087CW(CW);
        AFilter := nil;
      end;  
    except
      Result := false;
    end;
  end;

begin
  try
    Result := '';
    
    if not FileExists(videoFile) then begin
      Result := 'û���ҵ���Ҫ���ŵ���Ƶ�ļ�����鿴�ļ��Ƿ���ڡ�';
      Exit;
    end;



    FilterGraph.GraphEdit := _IsDebugFilter;

    //if FilterGraph.Active then begin

      FilterGraph.Stop;
      FilterGraph.ClearGraph;

      //FilterGraph.ClearGraph;
      //FilterGraph.Stop;   //�÷���ʹ��MediaControl����ֹͣ����

      FilterGraph.Active := False;


      VideoWindow.FilterGraph := nil;
      ImgCapture.FilterGraph := nil;

    //end;

    FilterGraph.Active := True;

    
    //��Ҫ��֤����vmr��windowlessģʽ�£��������samplegrabber�ӿڶ��󣬿����������������������Ƶ���ţ�
    //ʹ��sampleGrabberʱ����Щ��ʽ���ܽ���ץȡ
    //ImgCapture.FilterGraph := FilterGraph;  //�ڶ�ȡ�����ļ�֮ǰ����Ҫ���ʹ�õĽӿڣ��Ա��Զ���������

    //����ץȡ��ʽ������ʾģʽ(����ڲ�֧��VMR��ģʽ��ʹ��VMR�ǻ����һЩ���⣩
    if (_SnatchWay = swVMR) and (CheckIsSupportVmr) then begin
      VideoWindow.Mode := vmVMR;
      VideoWindow.VMROptions.Mode := vmrWindowless;

      VideoWindow.FilterGraph := FilterGraph;

      hr := (VideoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmr9windLessCtrl);
      if Succeeded(hr) then begin
        vwRect := Rect(0,0,VideoWindow.Width, VideoWindow.Height);

        vmr9windLessCtrl.SetVideoClippingWindow(VideoWindow.Handle);
        vmr9windLessCtrl.SetVideoPosition(nil, @vwRect);
      end;

    end else begin
      VideoWindow.Mode := vmNormal;
      VideoWindow.VMROptions.Mode := vmrWindowless;

      VideoWindow.FilterGraph := FilterGraph;

      hr := (VideoWindow as IBaseFilter).QueryInterface(IID_IVideoWindow, vw);
      if Succeeded(hr) then begin
        vw.put_Owner(VideoWindow.Handle);
        vw.put_WindowStyle(WS_CHILD or WS_CLIPCHILDREN or WS_CLIPSIBLINGS);
        vw.SetWindowPosition(0, 0, VideoWindow.Width, VideoWindow.Height);
      end;
    end;


    //��ȡ�����ļ����ڶ�ȡ�����ļ�֮ǰ����Ҫ���ʹ�õĽӿڣ��Ա��Զ���������
    //hr := FilterGraph.RenderFile(videoFile);
    try
    
      //��Ҫ��֤����vmr��windowlessģʽ�£��������samplegrabber�ӿڶ��󣬿����������������������Ƶ���ţ�
      //ʹ��sampleGrabberʱ����Щ��ʽ���ܽ���ץȡ
      ImgCapture.FilterGraph := FilterGraph;  //�ڶ�ȡ�����ļ�֮ǰ����Ҫ���ʹ�õĽӿڣ��Ա��Զ���������

      //hr := FilterGraph.RenderFile(videoFile);
      hr := FilterGraph.RenderFileEx(videoFile);
    except
      ImgCapture.FilterGraph := nil;
      
      //wmv����asf��ý���ļ�����ʹ��RenderFileEx�������в���
      hr := FilterGraph.RenderFile(videoFile);
    end;

    if Failed(hr) then begin
      Result := '�ļ���ȡ����[�������:' + IntToStr(hr) + ']';
      Exit;
    end;



    //�ȴ�RenderFileEx������ȫִ����ɣ��������InitVideoInf����ʱ�������Ҳ���ĳЩ�ӿ�
    Sleep(150);

    //��ȡ��Ƶ������Ϣ,
    InitVideoInf(videoFile);

    _VideoFile := videoFile;

    FilterGraph.Play;
    
    //��vb�е���ʱ����Ҫ���ϸþ䣬��������ʾ(form.show)���Ŵ��ں󣬽�����������ʱ������쳣
    Sleep(150);

    imgLogo.Visible := False;

    AdjustWindowSize();

    _VideoState := vsPlay;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.Stop: WideString;
begin
  try
    Result := '';

    //�˳�ȫ��ģʽ
    QuitFullScreen();


    FilterGraph.Stop;
    FilterGraph.ClearGraph;
    
    //FilterGraph.ClearGraph;
    //FilterGraph.Stop;   //�÷���ʹ��MediaControl����ֹͣ����

    FilterGraph.Active := False;

    _VideoState := vsStop;
    imgLogo.Visible := True;

    //�÷����ᴥ��videowindow��paint�¼�
    VideoWindow.Refresh;

    if Assigned(_VideoInf) then FreeAndNil(_VideoInf);
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.AddRate: WideString;
var
  ms: IMediaSeeking;
  hr: HRESULT;
  curRate: Double;
  newpos, curpos, stopPos: Int64;
begin
  try
    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����������Ƶ�Ĳ������ʡ�';
      Exit;
    end;

    try
      //ע���ڲ�����h264�ȱ����������һЩ��Ƶʱ�����������ʺ���Ҫִ��һ�ζ�λ�����ܼ������š�
      hr := ms.GetCurrentPosition(curPos);
      if Failed(hr) then Exit;

      hr := ms.GetStopPosition(stopPos);
      if Failed(hr) then Exit;

      hr := ms.GetRate(curRate);
      if Failed(hr) then Exit;
      
      ms.SetRate(curRate + _rateStep);
      if Failed(hr) then Exit;

      newpos := curPos + Trunc(ONE_SECOND / _VideoInf.FrameRate + 0.5);
      if newpos > stopPos then begin
        ms.SetPositions(stopPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
      end else begin
        ms.SetPositions(newpos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
      end;
    finally
      ms := nil;
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.CaptureImageToBmpObj: TBitmap;
var
  curBitMap: TBitmap;
  bmpStream: TMemoryStream;
begin
  if not Assigned(_VideoInf) or (_VideoState = vsStop) then begin
    Result := nil;
    Exit;
  end;

  if _VideoInf.MajorTypeName <> 'Video' then begin
    Result := nil;
    Exit;
  end;

  curBitMap := TBitmap.Create;
  try
    bmpStream := TMemoryStream.Create;
    try
      //�ɼ���֡ͼ��
      if (_SnatchWay = swVMR) then begin
        if (ImgCapture.FilterGraph = nil) or not ImgCapture.GetBitmap(curBitMap) then begin
          //���û��ץȡ��ͼ����ʹ��vmr�ӿڻ�ȡ
          VideoWindow.VMRGetBitmap(bmpStream);
          curBitMap.LoadFromStream(bmpStream);
        end;  

        Result := curBitMap;

        if _IsSoundHint then PostMessage(Self.Handle, WM_BEEPMSG, 0, 0);
        Exit;
      end;

      if (_SnatchWay = swDEVICE) then begin
        if (ImgCapture.FilterGraph <> nil) and ImgCapture.GetBitmap(curBitMap) then begin
          Result := curBitMap;
        end else begin
          Result := nil;
        end;

        if _IsSoundHint then PostMessage(Self.Handle, WM_BEEPMSG, 0, 0);
        Exit;
      end;

      Result := nil;
    finally
      FreeAndNil(bmpStream);
    end;
  except
    on e: Exception do begin
      Result := nil;
      if Assigned(curBitMap) then FreeAndNil(curBitMap);

      TDebug.DebugMsg('TDSPlay', 'CaptureImageToBmpObj', e.Message);
    end;
  end;
end;

function TDSPlay.CaptureBmpImgToFile(
  const fileName: WideString): WideString;
var
  bmp: TBitmap;
begin
  try
    Result := '';

    if not Assigned(_VideoInf) or (_VideoState = vsStop) then begin
      Result := '������δ���ڲ���״̬�����ܲɼ���Ƶͼ��';
      Exit;
    end;

    bmp := CaptureImageToBmpObj();

    if not Assigned(bmp) then begin
      Result := 'û�вɼ�����Ƶͼ��';
      Exit;
    end;

    try
      bmp.SaveToFile(fileName);
    finally
      FreeAndNil(bmp);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.CaptureJpgImgToFile(const fileName: WideString;
  compressRate: SYSINT): WideString;
var
  bmp: TBitmap;
  jpg: TJPEGImage;
begin
  try
    Result := '';

    if not Assigned(_VideoInf) or (_VideoState = vsStop) then begin
      Result := '������δ���ڲ���״̬�����ܲɼ���Ƶͼ��';
      Exit;
    end;    

    bmp := CaptureImageToBmpObj;

    if not Assigned(bmp) then begin
      Result := 'û�вɼ�����Ƶͼ��';
      Exit;
    end;

    try
      //ת��ͼ�񲢱���Ϊ�ļ�
      jpg := TGraphicProcess.BmpConvertToJpg(bmp, compressRate);
      try
        if not Assigned(jpg) then begin
          Result := 'ͼ��ת��ʧ�ܡ�';
          Exit;
        end;

        jpg.SaveToFile(fileName);
      finally
        if Assigned(jpg) then FreeAndNil(jpg);
      end;
    finally
      FreeAndNil(bmp);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.DecRate: WideString;
var
  ms: IMediaSeeking;
  hr: HRESULT;
  curRate: Double;
  newpos, curPos, stopPos: Int64;
begin
  try
    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����������Ƶ�Ĳ������ʡ�';
      Exit;
    end;

    try
      //ע���ڲ�����h264�ȱ����������һЩ��Ƶʱ�����������ʺ���Ҫִ��һ�ζ�λ�����ܼ������š�
      hr := ms.GetCurrentPosition(curPos);
      if Failed(hr) then Exit;

      hr := ms.GetStopPosition(stopPos);
      if Failed(hr) then Exit;

      ms.GetRate(curRate);
      if Failed(hr) then Exit;

      if (curRate - _rateStep) >= 0 then begin
        ms.SetRate(curRate - _rateStep);
        if Failed(hr) then Exit;
      end;

      newpos := curPos + Trunc(ONE_SECOND / _VideoInf.FrameRate + 0.5);
      if newpos > stopPos then begin
        ms.SetPositions(stopPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
      end else begin
        ms.SetPositions(newpos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
      end;

    finally
      ms := nil;
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.RestoreRate: WideString;
var
  mediaSeeking: IMediaSeeking;
  hr: HRESULT;
begin
  try
    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, mediaSeeking);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����������Ƶ�Ĳ������ʡ�';
      Exit;
    end;

    mediaSeeking.SetRate(1);

    mediaSeeking := nil;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.ShowVideoInfo(parentHandle: SYSINT): WideString;
begin
  try
    Result := TfrmVideoInf.ShowVideoInf(parentHandle, _VideoInf);
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

procedure TDSPlay.FreeRes;
begin
  try
    ReInitDSPlay();
  except
  end;
end;

procedure TDSPlay.InitDSPlay;
begin
  _VideoState := vsStop;
  _rateStep := 0.2;
  _VideoInf := nil;
  _ShowModel := smFit;
  _IsAdjustWindowSize := False;

  _IsEscKeyQuitFullScreen := True;
  _IsDblClickQuitFullScreen := False;
  _IsClickQuitFullScreen := False;

  _VideoFile := '';

  _SnatchWay := swVMR;
  _IsSoundHint := False;
  _IsDebugFilter := False;

  VideoWindow.Color := Color;
end;

procedure TDSPlay.ReInitDSPlay;
begin
  //�˳�ȫ��ģʽ
  QuitFullScreen();

  if _VideoState <> vsStop then begin
    FilterGraph.Stop;
    FilterGraph.ClearGraph;

    FilterGraph.Active := False;
  end;

  //�ͷ���Ƶ��Ϣ����
  if Assigned(_VideoInf) then FreeAndNil(_VideoInf);
end;

function TDSPlay.Run: WideString;
begin
  try
    Result := '';

    //ֻ����ͣ��ʱ�򣬲��ܹ�ִ�м������ŵĲ���
    if _VideoState <> vsPause then Exit;

    FilterGraph.Play;

    _VideoState := vsPlay;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.FirstFrame: WideString;
var
  mediaSeeking: IMediaSeeking;
  curPos, stopPos: Int64;
  hr: HRESULT;
begin
  try
    Result := '';

    if not Assigned(_VideoInf) then begin
      Result := 'û��ȡ���������Ƶ������Ϣ��';
      Exit;
    end;    

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, mediaSeeking);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����ָ��֡��λ�á�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;
    
    curPos := 0;
    stopPos := 0;

    mediaSeeking.SetPositions(curPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);

    mediaSeeking := nil;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.LastFrame: WideString;
var
  mediaSeeking: IMediaSeeking;
  curPos, stopPos: Int64;
  hr: HRESULT;
begin
  try
    Result := '';

    if not Assigned(_VideoInf) then begin
      Result := 'û��ȡ���������Ƶ������Ϣ��';
      Exit;
    end;    

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, mediaSeeking);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����ָ��֡��λ�á�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    curPos := 0;
    stopPos := 0;
        
    mediaSeeking.GetStopPosition(curPos);

    mediaSeeking.SetPositions(curPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);

    mediaSeeking := nil;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.NextFrame: WideString;
var
  mediaSeeking: IMediaSeeking;
  newPos, curPos, stopPos: Int64;
  hr: HRESULT;  
begin
  try
    Result := '';

    if not Assigned(_VideoInf) then begin
      Result := 'û��ȡ���������Ƶ������Ϣ��';
      Exit;
    end;    

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, mediaSeeking);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����ָ��֡��λ�á�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    hr := mediaSeeking.GetCurrentPosition(curPos);
    if Failed(hr) then begin
      Result := '����ȡ����Ƶ�ĵ�ǰ����λ�á�';
      Exit;
    end;

    //������һ֡�����ʱ��
    newPos := Trunc(ONE_SECOND / _VideoInf.FrameRate + 0.5);
    stopPos := 0;

    hr := mediaSeeking.GetStopPosition(stopPos);
    if Failed(hr) then begin
      Result := '����ȡ����Ƶ�Ľ���λ�á�';
      Exit;
    end;

    if curPos + newPos >= stopPos then begin
      mediaSeeking.SetPositions(stopPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
    end else begin
      newPos := curPos + newPos;  //���þ���λ�ö�λ
      mediaSeeking.SetPositions(newPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
    end;  

    mediaSeeking := nil;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.PriorFrame: WideString;
var
  mediaSeeking: IMediaSeeking;
  curPos, stopPos: Int64;
  hr: HRESULT;
begin
  try
    Result := '';

    if not Assigned(_VideoInf) then begin
      Result := 'û��ȡ���������Ƶ������Ϣ��';
      Exit;
    end;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, mediaSeeking);
    if Failed(hr) then begin
      Result := '��ѯ�ӿ�MediaSeekingʱʧ�ܣ�����ָ��֡��λ�á�[�������:' + IntToStr(hr) + ']';
      Exit;
    end;

    try
      hr := mediaSeeking.GetCurrentPosition(curPos);
      if Failed(hr) then begin
        Result := '����ȡ����Ƶ�ĵ�ǰ����λ�á�';
        Exit;
      end;

      //������һ֡����Ӧ�Ĳ���ʱ��
      curPos := Trunc(curPos - ONE_SECOND / _VideoInf.FrameRate + 0.5);
      stopPos := 0;

      if curPos < 0 then Exit;

      mediaSeeking.SetPositions(curPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);

    finally
      mediaSeeking := nil;
    end;  
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

procedure TDSPlay.InitVideoInf(const fileName: WideString);
var
  md: IMediaDet;
  //iVmr9: IVMRWindowlessControl9;
  //ibVideo: IBasicVideo;
  hr: HRESULT;
  iStreamCount: Integer;
  i: Integer;
  bFoundStream: Boolean;
  mt: TAMMediaType;
  pvih: PVideoInfoHeader;
  frameRate: Double;
  ms: IMediaSeeking;
  //mp: IMediaPosition;
  stopPos, stopTime: Int64;
  timeFormat: TGUID;
  //sRect, dRect: TRect;
  //vWidth, vHeight: Integer;
begin
  if not FileExists(fileName) then Exit;

  hr := CoCreateInstance(CLSID_MediaDet, nil, CLSCTX_INPROC_SERVER, IID_IMediaDet, md);
  if Failed(hr) then Exit;

  try
    hr := md.put_Filename(fileName);
    if Failed(hr) then Exit;


    //ȡ���ļ�������������Ƶ�ļ���������Ƶ������Ƶ������Ϣ����������ɣ�
    hr := md.get_OutputStreams(iStreamCount);
    if Failed(hr) then Exit;


    bFoundStream := False;
    for i := 0 to iStreamCount - 1 do begin
      hr := md.put_CurrentStream(i);
      if Failed(hr) then Continue;

      hr := md.get_StreamMediaType(mt);
      if Failed(hr) then Continue;

      if GUIDToString(mt.majortype) = GUIDToString(MEDIATYPE_Audio) then begin
        //�������Ƶ���������������Ƶ�������ҵ���Ƶ��ʱ��ֹͣ��ѭ��
        bFoundStream := True;
        Continue;
      end;
      
      //�������Ƶ�����������ǰѭ��
      if GUIDToString(mt.majortype) = GUIDToString(MEDIATYPE_Video) then begin
        bFoundStream := True;
        Break;
      end;
    end;

    //���û���ҵ���Ƶ�������˳�
    if not bFoundStream then Exit;
    //���������Ƶ���������ǰ����(���ִ�и���䣬�����ܲ�����Ƶ�ļ�)
    //if GUIDToString(mt.formattype) <> GUIDToString(FORMAT_VideoInfo) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if Failed(hr) then Exit;

    ms.GetStopPosition(stopPos);   //}

    {��������ʹ��IMediaSeeking��������Ϣ����˲���Ҫʹ��IMediaPosition�ӿ�
    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaPosition, mp);
    if Failed(hr) then Exit;

    mp.get_StopTime(stopTime1);    //}

    //ת��Ϊ��    GetStopPosition��ȡ��ֵ�ĵ�λΪ10΢��
    stopTime := Round(stopPos / ONE_SECOND);

    md.get_FrameRate(frameRate);
    ms.GetTimeFormat(timeFormat);

    pvih := mt.pbFormat;

    if Assigned(_VideoInf) then FreeAndNil(_VideoInf);

    //������Ƶ��Ϣ
    _VideoInf := TVideoInf.Create;

    _VideoInf.videoFile := fileName;
    _VideoInf.MajorTypeName := TDS9Ex.GetMediaGuidName(mt.majortype);
    _VideoInf.SubTypeName := TDS9Ex.GetMediaGuidName(mt.subtype);
    _VideoInf.FormatTypeName := TDS9Ex.GetMediaGuidName(mt.formattype);
    _VideoInf.TimeFormatName := TDS9Ex.GetTimeFormatName(timeFormat);
    _VideoInf.VideoColorDepth := pvih^.bmiHeader.biBitCount;
    _VideoInf.VideoWidth := pvih^.bmiHeader.biWidth;
    _VideoInf.VideoHeight := pvih^.bmiHeader.biHeight;
    _VideoInf.StreamCount := iStreamCount;
    _VideoInf.FrameRate := frameRate;
    _VideoInf.TimeLen := stopTime;
    _VideoInf.FrameLen := Trunc(stopPos / ONE_SECOND * frameRate + 0.5);

    //���ҵ���Ƶ��֮��������ע�͵����ɲ�ʹ��  2011-01-04
    {if _SnatchWay = swVMR then begin
      //���ʹ��vmr��ʾģʽ����asf�߼�����ʽ����Ƶʱ������ʹ��pvih�ṹ��ȡ�ֱ��ʴ�С����Ҫʹ��vmr9�ӿڻ�ȡ
      hr :=  (VideoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, ivmr9);
      if Failed(hr) then exit;

      try
        iVmr9.GetVideoPosition(sRect, dRect);

        _VideoInf.VideoWidth := sRect.Right;
        _VideoInf.VideoHeight := sRect.Bottom;
      finally
        iVmr9 := nil;
      end;
    end else begin
      if GUIDToString(mt.formattype) <> GUIDToString(FORMAT_WaveFormatEx) then Exit;

      //���δʹ��vmr��ʾ��ʽ  ����ͬ����Ƶ����ʹ�ò�ͬ�ķ�����ȡ�ֱ��ʴ�Сʱ��ȡ�õķֱ��ʲ�һ�£�
      hr :=  (VideoWindow as IBaseFilter).QueryInterface(IID_IBasicVideo, ibVideo);
      if Failed(hr) then exit;

      try
        ibVideo.get_SourceWidth(vWidth);
        ibVideo.get_SourceHeight(vHeight);

        _VideoInf.VideoWidth := vWidth;
        _VideoInf.VideoHeight := vHeight;
      finally
        ibVideo := nil;
      end;
    end;//}

  finally
    md := nil;
  end;
end;

function TDSPlay.Get_CurFrame: SYSINT;
var
  ms: IMediaSeeking;
  hr: HRESULT;
  curPos: Int64;
begin
  try
    Result := -1;

    if not Assigned(_VideoInf) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if Failed(hr) then Exit;

    ms.GetCurrentPosition(curPos);

    Result := Trunc(curPos / ONE_SECOND * _VideoInf.FrameRate + 0.5);

    ms := nil;
  except
    Result := -1;
  end;
end;

function TDSPlay.Get_CurTime: SYSINT;
var
  ms: IMediaSeeking;
  hr: HRESULT;
  curPos: Int64;
begin
  try
    Result := -1;

    if not Assigned(_VideoInf) then begin
      Exit;
    end;  

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if Failed(hr) then Exit;

    ms.GetCurrentPosition(curPos);

    Result := Trunc(curPos / ONE_SECOND + 0.5);

    ms := nil;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'Get_CurTime', e.Message);
      Result := -1;
    end;

  end;
end;

function TDSPlay.Get_FrameLen: SYSINT;
begin
  Result := 0;

  if Assigned(_VideoInf) then
    Result := _VideoInf.FrameLen;
end;

function TDSPlay.Get_TimeLen: SYSINT;
begin
  Result := 0;

  if Assigned(_VideoInf) then
    Result := _VideoInf.TimeLen;
end;

procedure TDSPlay.Set_CurFrame(Value: SYSINT);
var
  //ms: IMediaSeeking;    {imediaseking��֧���Զ����ӿڣ�IMediaPosition֧���Զ����ӿ�}
  mp: IMediaPosition;
  hr: HRESULT;
  //curPos, stopPos: Int64;
begin
  if Value < 0 then Exit;
  if not Assigned(_VideoInf) then Exit;

  if Value >= _VideoInf.FrameLen then Exit;

  hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaPosition, mp{IID_IMediaSeeking, ms});
  if Failed(hr) then Exit;

  try
    {//����ָ������֡����Ӧ�Ĳ���ʱ��
    curPos := Trunc(value / _VideoInf.FrameRate * ONE_SECOND + 0.5);
    stopPos := 0;

    ms.SetPositions(curPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
    }
    mp.put_CurrentPosition(Trunc(Value / _VideoInf.FrameRate))
  finally
    //ms := nil;
    mp := nil;
  end;  
end;

procedure TDSPlay.Set_CurTime(Value: SYSINT);
var
  //ms: IMediaSeeking;    {imediaseking��֧���Զ����ӿڣ�IMediaPosition֧���Զ����ӿ�}
  mp: IMediaPosition;
  hr: HRESULT;
  //curPos, stopPos: Int64;
begin
  if Value < 0 then Exit;
  if not Assigned(_VideoInf) then Exit;

  if Value >= _VideoInf.TimeLen then Exit;

  hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaPosition, mp{IID_IMediaSeeking, ms});
  if Failed(hr) then Exit;

  try
    {if Failed( ms.IsFormatSupported(TIME_FORMAT_FRAME)) then begin   //��֧�ָö�λ��ʽ
      exit;
    end;

    curPos := Value * ONE_SECOND;
    stopPos := 0;

    hr := ms.GetStopPosition(stopPos);
    if Failed(hr) then begin
      Exit;
    end;

    ms.SetPositions(curPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
    }

    mp.put_CurrentPosition(Value);


  finally
    //ms := nil;
    mp := nil;
  end;
end;

function TDSPlay.Get_PlayRate: Double;
var
  ms: IMediaSeeking;
  hr: HRESULT;  
begin
  try
    Result := -1;

    if not Assigned(_VideoInf) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
    if Failed(hr) then Exit;

    try
      ms.GetRate(Result);
    finally
      ms := nil;
    end;
  except
    Result := -1;
  end;
end;

procedure TDSPlay.Set_PlayRate(Value: Double);
var
  ms: IMediaSeeking;
  hr: HRESULT;
  newpos, curPos, stopPos: Int64;
begin
  if Value < 0 then Exit;
  if not Assigned(_VideoInf) then Exit;

  hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IMediaSeeking, ms);
  if Failed(hr) then Exit;

  try
    //ע���ڲ�����h264�ȱ����������һЩ��Ƶʱ�����������ʺ���Ҫִ��һ�ζ�λ�����ܼ������š�
    hr := ms.GetCurrentPosition(curPos);
    if Failed(hr) then Exit;

    hr := ms.GetStopPosition(stopPos);
    if Failed(hr) then Exit;
                   
    hr := ms.SetRate(Value);
    if Failed(hr) then Exit;

    newpos := curPos + Trunc(ONE_SECOND / _VideoInf.FrameRate + 0.5);
    if newpos > stopPos then begin
      ms.SetPositions(stopPos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
    end else begin
      ms.SetPositions(newpos, AM_SEEKING_AbsolutePositioning, stopPos, AM_SEEKING_NoPositioning);
    end;
  finally
    ms := nil;
  end;
end;

function TDSPlay.Get_VideoState: TVideoState;
begin
  Result := _VideoState;
end;

function TDSPlay.Get_ShowModel: TShowModel;
begin
  Result := _ShowModel;
end;

procedure TDSPlay.Set_ShowModel(Value: TShowModel);
begin
  _ShowModel := Value;
  //AdjustWindowSize();
end;

function TDSPlay.Get_IsAdjustWindowSize: WordBool;
begin
  Result := _ShowModel = smWindAutoFit;
  AdjustWindowSize();
end;

function TDSPlay.Get_IsFit: WordBool;
begin
  Result := _ShowModel = smFit;
end;

function TDSPlay.Get_IsFullScreen: WordBool;
begin
  Result := TfrmFullScreen.GetFullScreenState();
end;

function TDSPlay.Get_IsShowState: WordBool;
begin
  Result := stabStates.Visible;    
end;

function TDSPlay.Get_IsStretch: WordBool;
begin
  Result := _ShowModel = smStretch;
end;

procedure TDSPlay.Set_IsAdjustWindowSize(Value: WordBool);
begin
  _IsAdjustWindowSize := Value;

  if (_IsAdjustWindowSize = False) and (_IsFit = False) and (_IsStretch = False) then _ShowModel := smNormal;
  if (_IsAdjustWindowSize = False) and (_IsFit = False) and (_IsStretch = True) then _ShowModel := smStretch;
  if (_IsAdjustWindowSize = False) and (_IsFit = True) and (_IsStretch = False) then _ShowModel := smFit;
  if (_IsAdjustWindowSize = False) and (_IsFit = True) and (_IsStretch = True) then _ShowModel := smFit;

  if Value then begin
    _ShowModel := smWindAutoFit;
    _IsFit := False;
    _IsStretch := False;
  end;

  //AdjustWindowSize;
end;

procedure TDSPlay.Set_IsFit(Value: WordBool);
begin
  _IsFit := Value;

  if (_IsFit = False) and (_IsStretch = False) and (_IsAdjustWindowSize = False) then _ShowModel := smNormal;
  if (_IsFit = False) and (_IsStretch = False) and (_IsAdjustWindowSize = True) then _ShowModel := smWindAutoFit;
  if (_IsFit = False) and (_IsStretch = True) and (_IsAdjustWindowSize = True) then _ShowModel := smStretch;
  if (_IsFit = False) and (_IsStretch = True) and (_IsAdjustWindowSize = False) then _ShowModel := smStretch;

  if _IsFit then begin
    _ShowModel := smFit;
    _IsStretch := False;
    _IsAdjustWindowSize := False;
  end;

  //AdjustWindowSize();
end;

procedure TDSPlay.Set_IsFullScreen(Value: WordBool);
begin
  if Value then begin
    ShowFullScreen(Handle, 0);
  end else begin
    QuitFullScreen();
  end;
end;

procedure TDSPlay.Set_IsShowState(Value: WordBool);
begin
  stabStates.Visible := Value;

  if Assigned(_VideoInf) then AdjustWindowSize();
end;

procedure TDSPlay.Set_IsStretch(Value: WordBool);
begin
  _IsStretch := Value;

  if (_IsStretch = False) and (_IsFit = False) and (_IsAdjustWindowSize = False) then _ShowModel := smNormal;
  if (_IsStretch = False) and (_IsFit = False) and (_IsAdjustWindowSize = True) then _ShowModel := smWindAutoFit;
  if (_IsStretch = False) and (_IsFit = True) and (_IsAdjustWindowSize = True) then _ShowModel := smFit;
  if (_IsStretch = False) and (_IsFit = True) and (_IsAdjustWindowSize = False) then _ShowModel := smFit;

  if _IsStretch then begin
    _ShowModel := smStretch;
    _IsFit := False;
    _IsAdjustWindowSize := False;
  end;

  //AdjustWindowSize();
end;

procedure TDSPlay.VideoWindowPaint(Sender: TObject);
var
  vmrWindCtrl9: IVMRWindowlessControl9;
  vw: IVideoWindow;
  hr: HRESULT;
  videoDc: HDC;
begin
  if imgLogo.Visible then begin
    imgLogo.Left := (VideoWindow.Width - imgLogo.Width) div 2;
    imgLogo.Top := (VideoWindow.Height - imgLogo.Height) div 2;
  end;

  if imgAnimate.Visible then begin
    imgAnimate.Left := (Self.Width - imgAnimate.Width) div 2;
    imgAnimate.Top := (Self.Height - imgAnimate.Height) div 2;
  end;
  
  //��������Ļ�������������ʱ���ڻָ���ʱ��ˢ����Ƶ��ʾ
  if not Assigned(_VideoInf) or (_VideoState = vsStop) then Exit;

  if _SnatchWay = swVMR then begin
    hr := (VideoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindCtrl9);
    if Failed(hr) then Exit;

    videoDc := GetDC(VideoWindow.Handle);
    try
      vmrWindCtrl9.RepaintVideo(VideoWindow.Handle, videoDc);
    finally
      vmrWindCtrl9 := nil;
      ReleaseDC(VideoWindow.Handle, videoDc);
    end;
  end else begin
    //device��ʽ��ˢ��
    hr := (VideoWindow as IBaseFilter).QueryInterface(IID_IVideoWindow, vw);
    if Failed(hr) then Exit;

    try
      //ˢ����Ƶ
      vw.put_Visible(True);
    finally
      vw := nil;
    end;
  end;
end;

function TDSPlay.QuitFullScreen: WideString;
begin
  Result := TfrmFullScreen.QuitFullScreen();
end;

function TDSPlay.ShowFullScreen(parentHandle,
  monitorIndex: Integer): WideString;
begin
  if not Assigned(_VideoInf) or (_VideoState = vsStop)then begin
    Result := '��δ���벥��״̬�����ܽ���ȫ������ģʽ��';
    Exit;
  end;

  Result := TfrmFullScreen.ShowFullScreen(VideoWindow, _ShowModel, monitorIndex);
end;

procedure TDSPlay.AdjustWindowSize;
const
  BORDER_WIDTH: Integer = 0;
var
  rate: Double;
  stateBarHeight: Integer;
begin
  if not Assigned(_VideoInf) then begin
    VideoWindow.Left := 0;
    VideoWindow.Top := 0;
    
    VideoWindow.Width := Self.Width;
    VideoWindow.Height := Self.Height;

    if imgLogo.Visible then begin
      imgLogo.Left := (VideoWindow.Width - imgLogo.Width) div 2;
      imgLogo.Top := (VideoWindow.Height - imgLogo.Height) div 2;
    end;

    Exit;
  end;
  
  //�жϵ�ǰ���ڴ�С�Ƿ���Ӧ�ɼ����ڴ�С
  {if _IsAdjustWindowSize then begin
    Self.Width := _VideoInf.VideoWidth + 2;
    Self.Height := _VideoInf.VideoHeight + stabStates.Height + 2;

    VideoWindow.Width := _VideoInf.VideoWidth;
    VideoWindow.Height := _VideoInf.VideoHeight;

    VideoWindow.Left := (Self.Width - VideoWindow.Width) div 2 - 1;

    //�ж��Ƿ���Ҫ��ȥ״̬���߶�
    stateBarHeight := 0;
    if stabStates.Visible then stateBarHeight := stabStates.Height + 2;

    if VideoWindow.Height > Self.Height - stateBarHeight then begin
      VideoWindow.Top := Self.Height - VideoWindow.Height - stateBarHeight - 2;
    end else begin
      VideoWindow.Top := (Self.Height - stateBarHeight - VideoWindow.Height - BORDER_WIDTH) div 2;
    end;

    Exit;
  end;}

  if (_VideoInf.VideoWidth <= 0)  or (_VideoInf.VideoHeight <= 0)  then begin
    _ShowModel := smStretch;
    exit;
  end;

  //������ʾģʽ���ͣ��������λ�ü���С
  case _ShowModel of
    smNormal: begin //---------------------------------------------------------

      VideoWindow.Align := alNone;

      VideoWindow.Width := _VideoInf.VideoWidth;
      VideoWindow.Height := _VideoInf.VideoHeight;

      VideoWindow.Left := (Self.Width - VideoWindow.Width) div 2 - 1;

      //�ж��Ƿ���Ҫ��ȥ״̬���߶�
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height + 2;

      if _VideoInf.VideoHeight > Self.Height - stateBarHeight then begin
        VideoWindow.Top := Self.Height - VideoWindow.Height - stateBarHeight - 2;
      end else begin
        VideoWindow.Top := (Self.Height - stateBarHeight - VideoWindow.Height - BORDER_WIDTH) div 2;
      end;
    end;
    smFit: begin //-------------------------------------------------------------

      VideoWindow.Align := alNone;

      //�ж��Ƿ���Ҫ��ȥ״̬���߶�
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height + 2;

      //ȡ�����ű���
      if (_VideoInf.VideoHeight) / _VideoInf.VideoWidth > (Self.Height - stateBarHeight - BORDER_WIDTH) / (Self.Width - BORDER_WIDTH) then begin
        rate := (Self.Height - stateBarHeight - BORDER_WIDTH) / _VideoInf.VideoHeight;
      end else begin
        rate := Self.Width / _VideoInf.VideoWidth;
      end;

      //�����С��ȣ��򲻽�������
      if (_VideoInf.VideoHeight = Self.Height - stabStates.Height - 2)
        and (_VideoInf.VideoWidth = Self.Width - 2) then begin
        rate := 1;
      end;

      VideoWindow.Width := Round(_VideoInf.VideoWidth * rate);
      VideoWindow.Height := Round(_VideoInf.VideoHeight * rate);

      VideoWindow.Left := (Self.Width - VideoWindow.Width) div 2 - 1;
      VideoWindow.Top := (Self.Height - stateBarHeight - VideoWindow.Height - BORDER_WIDTH) div 2;
    end;
    smStretch: begin //---------------------------------------------------------

      VideoWindow.Align := alClient;
    end;
    smWindAutoFit: begin //-----------------------------------------------------
      //�жϵ�ǰ���ڴ�С�Ƿ���Ӧ�ɼ����ڴ�С

      //�ж��Ƿ���Ҫ��ȥ״̬���߶�
      stateBarHeight := 0;
      if stabStates.Visible then stateBarHeight := stabStates.Height;

      Self.Height := _VideoInf.VideoHeight + stateBarHeight + 2;
      Self.Width := _VideoInf.VideoWidth + 2;

      VideoWindow.Width := _VideoInf.VideoWidth;
      VideoWindow.Height := _VideoInf.VideoHeight;

      VideoWindow.Left := (Self.Width - VideoWindow.Width) div 2 - 1;

      if VideoWindow.Height > Self.Height - stateBarHeight then begin
        VideoWindow.Top := Self.Height - VideoWindow.Height - stateBarHeight;
      end else begin
        VideoWindow.Top := (Self.Height - stateBarHeight - VideoWindow.Height) div 2 - 1;
      end;
    end;
  end;
end;

function TDSPlay.RefreshWindow: WideString;
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

function TDSPlay.Get_IsClickQuitFullScreen: WordBool;
begin
  Result := _IsClickQuitFullScreen; 
end;

function TDSPlay.Get_IsDblClickQuitFullScreen: WordBool;
begin
  Result := _IsDblClickQuitFullScreen;
end;

function TDSPlay.Get_IsEscKeyQuitFullScreen: WordBool;
begin
  Result := _IsEscKeyQuitFullScreen;
end;

procedure TDSPlay.Set_IsClickQuitFullScreen(Value: WordBool);
begin
  _IsClickQuitFullScreen := Value;
end;

procedure TDSPlay.Set_IsDblClickQuitFullScreen(Value: WordBool);
begin
  _IsDblClickQuitFullScreen := Value;
end;

procedure TDSPlay.Set_IsEscKeyQuitFullScreen(Value: WordBool);
begin
  _IsEscKeyQuitFullScreen := Value;
end;

procedure TDSPlay.VideoWindowKeyDown(Sender: TObject; var Key: Word;
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

procedure TDSPlay.VideoWindowClick(Sender: TObject);
begin
  //�˳�ȫ��
  if _IsClickQuitFullScreen {and VideoWindow.FullScreen} then begin
    QuitFullScreen();
  end;      

  if Assigned(Self.OnClick) then Self.OnClick(Sender);
end;

procedure TDSPlay.VideoWindowDblClick(Sender: TObject);
begin
  //�˳�ȫ��
  if _IsDblClickQuitFullScreen {and VideoWindow.FullScreen} then begin
    QuitFullScreen();
  end;

  if Assigned(Self.OnDblClick) then Self.OnDblClick(Sender);
end;

procedure TDSPlay.VideoWindowEnter(Sender: TObject);
begin
  if Assigned(Self.OnEnter) then Self.OnEnter(Sender);
end;

procedure TDSPlay.VideoWindowExit(Sender: TObject);
begin
  if Assigned(Self.OnExit) then Self.OnExit(Sender);
end;

procedure TDSPlay.VideoWindowKeyPress(Sender: TObject; var Key: Char);
begin
  if Assigned(Self.OnKeyPress) then Self.OnKeyPress(Sender, Key);
end;

procedure TDSPlay.VideoWindowKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Assigned(Self.OnKeyUp) then Self.OnKeyUp(Sender, Key, Shift);
end;

procedure TDSPlay.VideoWindowMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(Self.OnMouseDown) then Self.OnMouseDown(Sender, Button, Shift, X, Y);
end;

procedure TDSPlay.VideoWindowMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  if Assigned(Self.OnMouseMove) then Self.OnMouseMove(Sender, Shift, x, y);
end;

procedure TDSPlay.VideoWindowMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(Self.OnMouseUp) then Self.OnMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TDSPlay.VideoWindowMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
begin
  if Assigned(Self.OnMouseWheel) then Self.OnMouseWheel(Sender, Shift, WheelDelta, MousePos, Handled);
end;

procedure TDSPlay.VideoWindowMouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if Assigned(Self.OnMouseWheelDown) then Self.OnMouseWheelDown(Sender, Shift, MousePos, Handled);
end;

procedure TDSPlay.VideoWindowMouseWheelUp(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  if Assigned(Self.OnMouseWheelUp) then Self.OnMouseWheelUp(Sender, Shift, MousePos, Handled);
end;

procedure TDSPlay.EnterEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnGotFocus;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'EnterEvent', e.Message);
    end;
  end;  
end;

procedure TDSPlay.ExitEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnLostFocus;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'ExitEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.KeyDownEvent(Sender: TObject; var Key: Word;
  Shift: TShiftState);
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
      TDebug.DebugMsg('TDSPlay', 'KeyDownEvent', e.Message);
    end;
  end;
end;

procedure TDSPlay.KeyUpEvent(Sender: TObject; var Key: Word;
  Shift: TShiftState);
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
      TDebug.DebugMsg('TDSPlay', 'KeyUpEvent', e.Message);
    end;
  end;     
end;

procedure TDSPlay.MouseDownEvent(Sender: TObject; Button: TMouseButton;
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
      TDebug.DebugMsg('TDSPlay', 'MouseDownEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.MouseMoveEvent(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
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
      TDebug.DebugMsg('TDSPlay', 'MouseMoveEvent', e.Message);
    end;
  end;     
end;

procedure TDSPlay.MouseUpEvent(Sender: TObject; Button: TMouseButton;
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
      TDebug.DebugMsg('TDSPlay', 'MouseUpEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.MouseWheelDownEvent(Sender: TObject; Shift: TShiftState;
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
      FEvents.OnMouseWheelDown(curShift, MousePos.X, MousePos.Y, curHandled);
      Handled := curHandled;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'MouseWheelDownEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.MouseWheelEvent(Sender: TObject; Shift: TShiftState;
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
      TDebug.DebugMsg('TDSPlay', 'MouseWheelEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.MouseWheelUpEvent(Sender: TObject; Shift: TShiftState;
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
      TDebug.DebugMsg('TDSPlay', 'MouseWheelUpEvent', e.Message);
    end;
  end;    
end;

procedure TDSPlay.ResizeEvent(Sender: TObject);
begin
  try
    if FEvents <> nil then FEvents.OnResize;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'ResizeEvent', e.Message);
    end;
  end;    
end;


procedure TDSPlay.timerStateTimer(Sender: TObject);
var
  curTime: Int64;
  stopTime: Int64;

  hour, min, sec, msec: Word;
  hour1, min1, sec1, msec1: Word;
begin
  try
    stabStates.Panels.Items[1].Text := TimeToStr(Now);

    case _VideoState of
      vsStop: begin
        stabStates.Panels.Items[3].Text := '��ֹͣ';
      end;
      vsPlay: begin
        stabStates.Panels.Items[3].Text := '������';
      end;
      vsPause: begin
        stabStates.Panels.Items[3].Text := '����ͣ';
      end;
    end;

    if _VideoState <> vsStop then begin
      curTime := int64(Get_CurTime) * int64(ONE_SECOND);
      stopTime := int64(Get_TimeLen) * int64(ONE_SECOND);

      DecodeTime(curTime div 10000 / MiliSecInOneDay, hour, min, sec, msec);
      DecodeTime(stopTime div 10000 / MiliSecInOneDay, hour1, min1, sec1, msec1);

      stabStates.Panels.Items[4].Text := Format('%d:%d:%d',[hour, min, sec]) + ' -- ' + Format('%d:%d:%d',[hour1, min1, sec1]);

      if (curTime >= stopTime) and (_VideoState = vsPlay) then begin
        Stop();
      end;
    end else begin
      stabStates.Panels.Items[4].Text := Format('%d:%d:%d',[0, 0, 0]) + ' -- ' + Format('%d:%d:%d',[0, 0, 0]);
    end;
      
  except
    on e: Exception do begin
      TDebug.DebugMsg('TDSPlay', 'timerStateTimer', e.Message);
    end;
  end;
end;

function TDSPlay.GetVideoProperty(propertyType: TVideoProperty;
  var value: WideString): WideString;
begin
  try
    if not Assigned(_VideoInf) then begin
      Result := 'û��ȡ���������Ƶ������Ϣ��';
      Exit;
    end;

    //������ص����Ե�����ֵ
    case propertyType of
      vpVideoFile: begin
        value := _VideoInf.videoFile;
      end;
      vpMajorTypeName: begin
        value := _VideoInf.MajorTypeName;
      end;
      vpSubTypeName: begin
        value := _VideoInf.SubTypeName;
      end;
      vpFormatTypeName: begin
        value := _VideoInf.FormatTypeName;
      end;
      vpTimeFormatName: begin
        value := _VideoInf.TimeFormatName;
      end;
      vpVideoColorDepth: begin
        value := IntToStr(_VideoInf.VideoColorDepth);
      end;
      vpVideoWidth: begin
        value := IntToStr(_VideoInf.VideoWidth);
      end;
      vpVideoHeight: begin
        value := IntToStr(_VideoInf.VideoHeight);
      end;
      vpStreamCount: begin
        value := IntToStr(_VideoInf.StreamCount);
      end;
      vpFrameRate: begin
        value := FloatToStr(_VideoInf.FrameRate);
      end;
      vpTimeLen: begin
        value := IntToStr(_VideoInf.TimeLen);
      end;
      vpFrameLen: begin
        value := IntToStr(_VideoInf.FrameLen);
      end;
    end;
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.RePlay: WideString;
begin
  Result := Play(_VideoFile);
end;

function TDSPlay.Get_CurHeight: Integer;
begin
  Result := Self.Height;
end;

function TDSPlay.Get_CurWidth: Integer;
begin
  Result := Self.Width;
end;

procedure TDSPlay.Set_CurHeight(Value: Integer);
begin
  Self.Height := Value;
end;

procedure TDSPlay.Set_CurWidth(Value: Integer);
begin
  Self.Width := Value;
end;

function TDSPlay.Get_SnatchWay: TSnatchWay;
begin
  Result := _SnatchWay;
end;

procedure TDSPlay.Set_SnatchWay(Value: TSnatchWay);
begin
  if _SnatchWay = Value then Exit;
  //ʹ��������Ч��������ڲ���״̬�������½��в���
  _SnatchWay := Value;

  if _VideoState <> vsStop then begin
    Play(_VideoFile);
  end;
end;

function TDSPlay.Get_AppHandle: Integer;
begin
  Result := Application.Handle;
end;

procedure TDSPlay.Set_AppHandle(Value: Integer);
begin
  Application.Handle := Value;
end;

function TDSPlay.CaptureImgToClipBoard: WideString;
var
  bitMap: TBitmap;
  data: THandle;
  palette :HPALETTE;
  curFormat: Word;
  //H: THandle;
  //P: Pointer;
begin
  try
    Result := '';

    if not Assigned(_VideoInf) or (_VideoState = vsStop) then begin
      Result := '������δ���ڲ���״̬�����ܲɼ���Ƶͼ��';
      Exit;
    end;    

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := 'û�вɼ�����Ƶͼ��';
        Exit;
      end;


      //ֱ�ӽ�ͼ���Ƶ�������
      bitMap.SaveToClipboardFormat(curFormat, data, palette);
      Clipboard.SetAsHandle(curFormat, data);

    finally
      FreeAndNil(bitMap);
    end;          
  except
    on e: Exception do begin
      Result := e.Message;
    end;
  end;
end;

function TDSPlay.Get_Balance: Integer;
var
  iAudio: IBasicAudio;    //a number from �C10,000 to 10,000
  hr: HRESULT;
begin
  try
    Result := 0;

    if not Assigned(_VideoInf) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IBasicAudio, iAudio);
    if Failed(hr) then Exit;

    iAudio.get_Balance(Result);

    iAudio := nil;
  except
    Result := 0;
  end;
end;

function TDSPlay.Get_Volume: Integer;
var
  iAudio: IBasicAudio;    //a number from �C10,000 to 0, 0��ʾ�������
  hr: HRESULT;
begin
  try
    Result := 0;

    if not Assigned(_VideoInf) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IBasicAudio, iAudio);
    if Failed(hr) then Exit;

    iAudio.get_Volume(Result);

    //��-10000��0��ֵת��Ϊ0��10000
    Result := 10000 + Result;

    iAudio := nil;
  except
    Result := 0;
  end;
end;

procedure TDSPlay.Set_Balance(Value: Integer);
var
  iAudio: IBasicAudio;    //a number from �C10,000 to 10,000
  hr: HRESULT;
begin
  try
    if not Assigned(_VideoInf) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IBasicAudio, iAudio);
    if Failed(hr) then Exit;

    iAudio.put_Balance(Value);    
  finally
    iAudio := nil;
  end;
end;

procedure TDSPlay.Set_Volume(Value: Integer);
var
  iAudio: IBasicAudio;    //a number from �C10,000 to 0, 0��ʾ�������
  hr: HRESULT;
begin
  try
    if not Assigned(_VideoInf) then Exit;

    hr := (FilterGraph as IGraphBuilder).QueryInterface(IID_IBasicAudio, iAudio);
    if Failed(hr) then Exit;

    //��0��10000��ֵת��Ϊ-10000��0
    iAudio.put_Volume(Value - 10000 );    
  finally
    iAudio := nil;
  end;
end;

function TDSPlay.Get_StreamTypeName: WideString;
begin
  Result := _VideoInf.MajorTypeName;
end;

procedure TDSPlay.ShowAnimate(AnimateType: TAnimateType);
{var
  gif: TGIFImage;  

  function GetRes(ResName: string): String;
  var
    resObj: TResourceStream ;
  begin
    Result := '';

    if FileExists('c:\Temp\' + ResName + '.gif') then begin
      Result :=  'c:\Temp\' + ResName + '.gif';
      ShowMessage(Result);
      exit;
    end;

    resObj := TResourceStream.Create(Handle, ResName, 'GIF');
    try
      Result := 'c:\Temp\' + ResName + '.gif';
      resObj.SaveToFile(Result);
    finally
      FreeAndNil(resObj);
    end;
  end; }
begin
  {try
  imgAnimate.Visible := true;
  gif := TGIFImage.Create;
  try
    case AnimateType of
      atQiu: gif.LoadFromFile(GetRes('midi'));
      atMidi: gif.LoadFromFile(GetRes('midi'));
      atLogon: begin
        imgAnimate.Visible := false;
        imgLogo.Visible := true;

        exit;
      end;
    end;

    gif.Animate := true;

    imgAnimate.Picture.Assign(gif);

    imgAnimate.Left := (Self.Width - imgAnimate.Width) div 2;
    imgAnimate.Height := (Self.Height - imgAnimate.Height) div 2;

    if imgLogo.Visible then imgLogo.Visible := false;


  finally
    FreeAndNil(gif);
  end;
  except
    on e: exception do begin
      ShowMessage(e.Message );
    end;
  end; }

end;

procedure TDSPlay.HideAnimate;
begin
  //imgAnimate.Visible := false;
end;

function TDSPlay.CaptureBmpImage: IPictureDisp;
var
  bitMap: TBitmap;
  picImage: TPicture;
begin
  try
    Result := nil;

    if not Assigned(_VideoInf) or (_VideoState = vsStop) then begin
      Exit;
    end;

    bitMap := CaptureImageToBmpObj;
    try
      if not Assigned(bitMap) then begin
        Result := nil;
        Exit;
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

function TDSPlay.Get_IsSoundHint: WordBool;
begin
  Result := _IsSoundHint
end;

procedure TDSPlay.Set_IsSoundHint(Value: WordBool);
begin
  _IsSoundHint := Value;
end;

procedure TDSPlay.WM_BEEP(var msg: TMessage);
begin
  Windows.Beep(2000, 500);
end;

function TDSPlay.Get_IsDebugFilter: WordBool;
begin
  Result := _IsDebugFilter;
end;

procedure TDSPlay.Set_IsDebugFilter(Value: WordBool);
begin
  _IsDebugFilter := Value;
end;

function TDSPlay.Get_VideoFile: WideString;
begin
  Result := _VideoFile;
end;

procedure TDSPlay.Set_VideoFile(const Value: WideString);
begin
  _VideoFile := Value;
end;

initialization
  TActiveFormFactory.Create(
    ComServer,
    TActiveFormControl,
    TDSPlay,
    Class_DSPlay,
    2,
    '',
    OLEMISC_SIMPLEFRAME or OLEMISC_ACTSLIKELABEL,
    tmApartment);
end.
