{*******************************************************************************
采集参数设置
创建人：TJH
创建日前：2009-11-3

描述：...

当 Filter连接FilterGraphic之后，Filter可直接转换为IBaseFilter接口
否则只有通过Filter.BaseFilter.CreateFilter取得IBaseFilter接口

*******************************************************************************}

unit CapParameterCfg;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, ComCtrls, DSPack, DSUtil, DirectShow9,
  IniFiles, VideoProcessDefine, jpeg, SizerControl, ZLDSVideoProcess_TLB, CaptureDebug;

type
  TfrmCapParameterCfg = class(TForm)
    pgcParameters: TPageControl;
    tsDevice: TTabSheet;
    pnlControl: TPanel;
    btnSure: TBitBtn;
    pnlDescription: TPanel;
    btnCancel: TButton;
    tsQuality: TTabSheet;
    tsVideo: TTabSheet;
    bvl1: TBevel;
    grp4: TGroupBox;
    lblVideoBrightness: TLabel;
    lblVideoContrast: TLabel;
    lblVideoHue: TLabel;
    lblVideoSaturation: TLabel;
    trckbrVideoBrightness: TTrackBar;
    trckbrVideoContrast: TTrackBar;
    trckbrVideoHue: TTrackBar;
    trckbrVideoSaturation: TTrackBar;
    btnDefault: TButton;
    lblBrightnessValue: TLabel;
    lblContrast: TLabel;
    lblHue: TLabel;
    lblSaturation: TLabel;
    grp6: TGroupBox;
    lblSec: TLabel;
    lbl3: TLabel;
    edtVideoCaptureTimes: TEdit;
    chkTimeLimit: TCheckBox;
    lblGamma: TLabel;
    lblGammaValue: TLabel;
    lblCaptureDevice: TLabel;
    lblAnalogVideo1: TLabel;
    cbbCaptureDevice: TComboBox;
    cbbAnalogVideo: TComboBox;
    trckbrVideoWhiteBlance: TTrackBar;
    lblWhiteBlance: TLabel;
    lblWhiteBlanceValue: TLabel;
    lbl5: TLabel;
    cbbVideoEncoder: TComboBox;
    lblVideoEncoder: TLabel;
    trckbrVideoGamma: TTrackBar;
    tsImageCapture: TTabSheet;
    panCut: TPanel;
    imgCapture: TImage;
    Image1: TImage;
    labRightValue: TLabel;
    labDownValue: TLabel;
    labPortrait: TLabel;
    cbApplyImgCut: TCheckBox;
    labPlane: TLabel;
    shapArea: TShape;
    Button1: TButton;
    tsVfwConfig: TTabSheet;
    btnVideoSourceProperty: TButton;
    btnVideoPinProperty: TButton;
    btnVideoFormatProperty: TButton;
    btnVideoDisplayProperty: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    rgpColorDepth: TRadioGroup;
    cbxVideoSize: TComboBox;
    butVideoEncoderProperty: TButton;
    chkConvertToGray: TCheckBox;
    tsVideoDisplay: TTabSheet;
    rdoGupShowModel: TRadioGroup;
    rdoGupSnatchWay: TRadioGroup;
    chkIsShowState: TCheckBox;
    butVideoCompressCfg: TButton;
    Label6: TLabel;
    cbxInput: TComboBox;
    Label7: TLabel;
    cbxOutput: TComboBox;
    butTest: TButton;
    Button2: TButton;
    Label8: TLabel;
    Label9: TLabel;
    chkIsAutoBrightness: TCheckBox;
    chkIsAutoContrast: TCheckBox;
    chkIsAutoHue: TCheckBox;
    chkIsAutoSaturation: TCheckBox;
    chkIsAutoGamma: TCheckBox;
    chkIsAutoWhiteBlance: TCheckBox;
    Label10: TLabel;
    cbSoundHint: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure cbbCaptureDeviceChange(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure trckbrVideoBrightnessChange(Sender: TObject);
    procedure trckbrVideoContrastChange(Sender: TObject);
    procedure trckbrVideoHueChange(Sender: TObject);
    procedure trckbrVideoSaturationChange(Sender: TObject);
    procedure btnDefaultClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure cbbAnalogVideoChange(Sender: TObject);
    procedure trckbrVideoGammaChange(Sender: TObject);
    procedure trckbrVideoWhiteBlanceChange(Sender: TObject);
    procedure cbbVideoEncoderChange(Sender: TObject);
    procedure chkTimeLimitClick(Sender: TObject);
    procedure edtVideoCaptureTimesChange(Sender: TObject);
    procedure chkConvertToGrayClick(Sender: TObject);
    procedure btnSureClick(Sender: TObject);
    procedure tsImageCaptureShow(Sender: TObject);
    procedure cbApplyImgCutClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnVideoSourcePropertyClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure pgcParametersChanging(Sender: TObject;
      var AllowChange: Boolean);
    procedure cbxVideoSizeChange(Sender: TObject);
    procedure rgpColorDepthClick(Sender: TObject);
    procedure butVideoEncoderPropertyClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure rdoGupSnatchWayClick(Sender: TObject);
    procedure rdoGupShowModelClick(Sender: TObject);
    procedure chkIsShowStateClick(Sender: TObject);
    procedure butTestClick(Sender: TObject);
    procedure cbxInputChange(Sender: TObject);
    procedure cbxOutputChange(Sender: TObject);
    procedure cbxVideoSizeDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure chkIsAutoBrightnessClick(Sender: TObject);
    procedure chkIsAutoContrastClick(Sender: TObject);
    procedure chkIsAutoHueClick(Sender: TObject);
    procedure chkIsAutoSaturationClick(Sender: TObject);
    procedure chkIsAutoGammaClick(Sender: TObject);
    procedure chkIsAutoWhiteBlanceClick(Sender: TObject);
    procedure cbSoundHintClick(Sender: TObject);
  private
    { Private declarations }
    _ICapGraphBuilder2: ICaptureGraphBuilder2;  //ICaptureGraphBuilder2
    _capSource: IBaseFilter;

    _IniFile: TIniFile;   //INI读写对象,用于存取一些参数选项

    _CurParameter: TCaptureParameter;  //当前采集配置参数
    _OldParameter: TCaptureParameter;  //原始采集参数保存

    _CaptureParameterCfgFileName: WideString; //采集参数的配置文件名称

    _OnParameterChangeEvent: TOnParameterChangeEvent;   //参数改变事件
    _OnVfwConfigCallEvent: TOnVfwConfigCallEvent; //vfw配置调用事件
    
    _IsAllowWriteCapturePar: Boolean;  //是否允许设置采集参数
    _SizeControl: TSizeControl;

    _IsSaveCapParameter: Boolean;        //是否保存采集参数

    _PositionType: TCapParameterPostion;

    _VideoFormats: String;               //保存取得的视频格式字符（视频分辨率大小）


    {***********************与设备读取设置相关函数******************************}


    //根据设备名称取得指定的FILTER
    function GetDeviceFilter(const deviceName: WideString): TFilter;

    //载入采集设备
    procedure LoadCaptureDevice();
    //载入编码器
    procedure LoadVideoEncoder();
    //从配置文件中载入视频制式
    procedure LoadVideoAnalogFromConfigFile();
    //根据配置文件载入视频大小
    procedure LoadVideoFormatFromConfigFile();
    //载入视频质量
    procedure LoadVideoQuality(const deviceName: WideString; const isLoadDefault: Boolean);
    //取得设备所支持的分辨率
    function GetVideoFormats(const deviceName: WideString): String;
    //载入视频端子
    //procedure LoadCrossbar();
    procedure LoadCrossbar1(const deviceName: WideString);
    procedure LoadCrossbar2(const deviceName: WideString);

    //取得实际的视频分辨率大小
    function GetRealVideoSize: TVideoSize;

    {***********************与参数读取保存相关函数******************************}

    //读取配置文件
    procedure ReadCaptureParameterCfgToFace();
    //设置参数对象值
    procedure SetParameterValue(const parameterType: TCaptureParameterType;
      const parameterValue: WideString; const refreshVideo: Boolean);


    {***********************与裁剪范围设置相关函数******************************}

    //载入采集图像
    procedure LoadCuptureImage();
    //载入裁剪范围
    procedure LoadImageCutArea();
    //设置图像调整范围
    procedure SetImageDefaultAdjustArea();
    //缩放图像
    procedure ZoomImageSize();
    //裁剪范围调整事件
    procedure ImageCutAreaChangeEvent(Sender: TObject; ControlRect: TRect);

  public
    { Public declarations }

    //显示视频参数配置对话框
    class procedure ShowCaptureParameterCfg(
      capGraphBuilder2: ICaptureGraphBuilder2;
      capSource: IBaseFilter;
      const cfgFileName: WideString;
      const capParameter: TCaptureParameter;
      const hideCfgItem: Integer;
      const postionType: TCapParameterPostion;
      const parentHandle: HWND;
      callBack: TOnParameterChangeEvent;
      vfwConfigCall: TOnVfwConfigCallEvent);
      
    //从文件创建采集的参数
    class procedure ReadCaptureParameterFromFile(
      const filename: WideString; var captureParameter: TCaptureParameter; var otherPar: TOtherPar);
    //保存采集参数
    class procedure WriteCaptureParameterToFile(
      const filename: WideString; captureParameter: TCaptureParameter);

    //取得采集的样品文件
    class function GetCaptureSampleFile(): WideString;


    //初始化界面参数
    procedure InitParameterCfg(const cfgFileName: WideString;
      const capParameter: TCaptureParameter);

    //隐藏指定的参数配置项目  
    procedure HideParameterCfgItem(const hideCfgItem: Integer);
      
    //窗口显示位置
    property PositionType: TCapParameterPostion read _PositionType write _PositionType;    
    //参数改变事件
    property OnParameterChange: TOnParameterChangeEvent read _OnParameterChangeEvent write _OnParameterChangeEvent;
    //vfw配置调用事件
    property OnVfwConfigCall: TOnVfwConfigCallEvent read _OnVfwConfigCallEvent write _OnVfwConfigCallEvent;

    property CapGraphBuilder2: ICaptureGraphBuilder2 read _ICapGraphBuilder2 write _ICapGraphBuilder2;

    property CapSourceFilter: IBaseFilter read _capSource write _capSource;
  end;



  
implementation

uses Types, DirectShow9Ex, ComObj, ACTIVEX, Math;

{$R *.dfm}

const
  CaptureCfgFileName: String = 'CaptureConfig.ini';  //采集参数配置文件名

  Section_ParameterCfg: String = 'CaptureParameter';  //参数配置节
  Section_VideoAnalogCfg: String = 'VideoAnalog';     //视频制式配置节
  Section_VideoFormatCfg: String = 'VideoFormat';     //视频制式配置节  

  AdjustImageSize: Integer = 270;

{ TfrmCapParameterCfg }

procedure TfrmCapParameterCfg.ReadCaptureParameterCfgToFace();
var
  findIndex: Integer;
begin
  //设备名称
  if Trim(_CurParameter.CaptureDeviceName) <> '' then begin
    //如果从配置读取的设备与当前默认的设备不同，则重新读取视频质量和采集端口
    findIndex := cbbCaptureDevice.Items.IndexOf(_CurParameter.CaptureDeviceName);
    if findIndex < 0 then begin
      Application.MessageBox(PChar('无效的采集设备 [' + PWideChar(_CurParameter.CaptureDeviceName) + ']，请重新设置。'), '提示', MB_OK + MB_ICONINFORMATION);
      exit;
    end;

    if (cbbCaptureDevice.ItemIndex <> findIndex)then begin
      cbbCaptureDevice.ItemIndex := findIndex;
      //根据采集设备载入质量信息
      LoadVideoQuality(cbbCaptureDevice.Text, False);
      //根据具体的设备配置，载入采集端口
      LoadCrossbar1(cbbCaptureDevice.Text);
    end;
  end else begin
    _CurParameter.CaptureDeviceName := cbbCaptureDevice.Text;
  end;  

  if Trim(cbbCaptureDevice.Text) = '' then Exit;

  
  //视频制式  
  if Trim(_CurParameter.VideoAnalog) <> '' then
    cbbAnalogVideo.ItemIndex := cbbAnalogVideo.Items.IndexOf(_CurParameter.VideoAnalog)
  else
    _CurParameter.VideoAnalog := cbbAnalogVideo.Text;

  //视频分辨率
  if Trim(_CurParameter.VideoSize) <> '' then  begin
    cbxVideoSize.ItemIndex := cbxVideoSize.Items.IndexOf(_CurParameter.VideoSize);
  end else begin
    _CurParameter.VideoSize := cbxVideoSize.Text;
  end;

  //颜色深度
  if _CurParameter.ColorDepth > 0 then begin
    case _CurParameter.ColorDepth of
      8: rgpColorDepth.ItemIndex := 0;
      24: rgpColorDepth.ItemIndex := 1;
      12: rgpColorDepth.ItemIndex := 2;
      32: rgpColorDepth.ItemIndex := 3;
      16: rgpColorDepth.ItemIndex := 4;
    end;
  end else begin
    case rgpColorDepth.ItemIndex of
      0: _CurParameter.ColorDepth := 8;
      1: _CurParameter.ColorDepth := 24;
      2: _CurParameter.ColorDepth := 12;
      3: _CurParameter.ColorDepth := 32;
      4: _CurParameter.ColorDepth := 16;
    end;
  end;


  //读取采集输入端口
  if _CurParameter.InputCrossbar >= 0 then
    cbxInput.ItemIndex := _CurParameter.InputCrossbar
  else
    _CurParameter.InputCrossbar := cbxInput.ItemIndex;

  //读取采集输出端口
  if _CurParameter.OutputCrossbar >= 0 then
    cbxOutput.ItemIndex := _CurParameter.OutputCrossbar
  else
    _CurParameter.OutputCrossbar := cbxOutput.ItemIndex;


  //亮度
  if _CurParameter.Brightness > 0 then
    trckbrVideoBrightness.Position := _CurParameter.Brightness
  else
    _CurParameter.Brightness := trckbrVideoBrightness.Position;

  //对比度  
  if _CurParameter.Contrast > 0 then
    trckbrVideoContrast.Position := _CurParameter.Contrast
  else
    _CurParameter.Contrast := trckbrVideoContrast.Position;
        
  //色调
  if _CurParameter.Hue > 0 then
    trckbrVideoHue.Position := _CurParameter.Hue
  else
    _CurParameter.Hue := trckbrVideoHue.Position;

  //饱和度  
  if _CurParameter.Saturation > 0 then
    trckbrVideoSaturation.Position := _CurParameter.Saturation
  else
    _CurParameter.Saturation := trckbrVideoSaturation.Position;

  //伽马
  if _CurParameter.Gamma > 0 then
    trckbrVideoGamma.Position := _CurParameter.Gamma
  else
    _CurParameter.Gamma := trckbrVideoGamma.Position;

  //白平衡
  if _CurParameter.WhiteBlance > 0 then
    trckbrVideoWhiteBlance.Position := _CurParameter.WhiteBlance
  else
    _CurParameter.WhiteBlance := trckbrVideoWhiteBlance.Position;

  //编码器
  if Trim(_CurParameter.EncoderName) <> '' then
    cbbVideoEncoder.ItemIndex := cbbVideoEncoder.Items.IndexOf(_CurParameter.EncoderName)
  else
    _CurParameter.EncoderName := cbbVideoEncoder.Text;

  //是否限时模式  
  chkTimeLimit.Checked := _CurParameter.IsTimeLimit;

  //限时时长
  edtVideoCaptureTimes.Text := IntToStr(_CurParameter.LimitLength);

  //是否转换为8位图
  chkConvertToGray.Checked := _CurParameter.IsConvertGrayImg;

  //显示模式
  rdoGupShowModel.ItemIndex := _CurParameter.VideoShowModel;

  //抓取方式
  rdoGupSnatchWay.ItemIndex := _CurParameter.SnatchWay;

  //是否显示视频状态
  chkIsShowState.Checked := _CurParameter.IsShowState;

  //视频质量自动
  chkIsAutoBrightness.Checked  := _CurParameter.IsAutoBrightness;
  chkIsAutoContrast.Checked    := _CurParameter.IsAutoContrast;
  chkIsAutoHue.Checked         := _CurParameter.IsAutoHue;
  chkIsAutoSaturation.Checked  := _CurParameter.IsAutoGamma;
  chkIsAutoGamma.Checked       := _CurParameter.IsAutoSaturation;
  chkIsAutoWhiteBlance.Checked := _CurParameter.IsAutoWhiteBlance;

  //是否进行声音提示
  cbSoundHint.Checked          := _CurParameter.IsSoundHint;
end;

procedure TfrmCapParameterCfg.SetParameterValue(
  const parameterType: TCaptureParameterType;
  const parameterValue: WideString; const refreshVideo: Boolean);
begin
  if not _IsAllowWriteCapturePar then Exit;
  
  case parameterType of
    cptCaptureDeviceName: begin      //采集设备
      _CurParameter.CaptureDeviceName := parameterValue;
    end;
    cptVideoAnalog: begin            //视频制式
      _CurParameter.VideoAnalog := parameterValue; 
    end;
    cptColorDepth: begin             //颜色深度
      _CurParameter.ColorDepth := StrToInt(parameterValue);
    end;
    cptVideoSize: begin              //分辨率
      _CurParameter.VideoSize := parameterValue;
    end;
    cptBrightness: begin             //亮度
      _CurParameter.Brightness := StrToInt(parameterValue);
    end;
    cptContrast: begin               //对比度
      _CurParameter.Contrast := StrToInt(parameterValue);
    end;
    cptHue: begin                    //色调
      _CurParameter.Hue := StrToInt(parameterValue);
    end;
    cptSaturation: begin             //饱和度
      _CurParameter.Saturation := StrToInt(parameterValue);
    end;
    cptGamma: begin                  //伽马
      _CurParameter.Gamma := StrToInt(parameterValue);
    end;
    cptWhiteBlance: begin            //白平衡
      _CurParameter.WhiteBlance := StrToInt(parameterValue);
    end;
    cptEncoderName: begin            //编码器名称
      _CurParameter.EncoderName := parameterValue;
    end;
    cptIsTimeLimit: begin            //是否时间限制
      _CurParameter.IsTimeLimit := StrToBool(parameterValue);
    end;
    cptLimitLength: begin            //时间限制长度
      _CurParameter.LimitLength := StrToInt(parameterValue);
    end;
    cptIsConvert8Bit: begin          //是否转换为8位
      _CurParameter.IsConvertGrayImg := StrToBool(parameterValue);
    end;
    cptIsApplyImageCut: begin        //是否应用裁剪设置
      _CurParameter.IsApplyImageCut := StrToBool(parameterValue);
    end;
    cptTopRate: begin                //top设置
      _CurParameter.TopRate := StrToFloat(parameterValue);
    end;
    cptHeightRate: begin             //height设置
      _CurParameter.HeightRate := StrToFloat(parameterValue);
    end;
    cptLeftRate: begin               //left设置
      _CurParameter.LeftRate := StrToFloat(parameterValue);
    end;
    cptWidthRate: begin              //width设置
      _CurParameter.WidthRate := StrToFloat(parameterValue);
    end;
    cptSnatchWay: begin              //SnatchWay设置
      _CurParameter.SnatchWay := StrToInt(parameterValue);
    end;
    cptVideoShowModel: begin         //VideoShowModel设置
      _CurParameter.VideoShowModel := StrToInt(parameterValue);
    end;
    cptIsShowState: begin            //IsShowState设置
      _CurParameter.IsShowState := StrToBool(parameterValue);
    end;
    cptInputCrossbar: begin
      _CurParameter.InputCrossbar := StrToInt(parameterValue);
    end;
    cptOutputCrossbar: begin
      _CurParameter.OutputCrossbar := StrToInt(parameterValue); 
    end;
    cptIsAutoBrightness: begin
      _CurParameter.IsAutoBrightness := StrToBool(parameterValue); 
    end;
    cptIsAutoContrast: begin
      _CurParameter.IsAutoContrast := StrToBool(parameterValue);
    end;
    cptIsAutoHue: begin
      _CurParameter.IsAutoHue := StrToBool(parameterValue);
    end;
    cptIsAutoSaturation: begin
      _CurParameter.IsAutoSaturation := StrToBool(parameterValue);
    end;
    cptIsAutoGamma: begin
      _CurParameter.IsAutoGamma := StrToBool(parameterValue);
    end;
    cptIsAutoWhiteBlance: begin
      _CurParameter.IsAutoWhiteBlance := StrToBool(parameterValue);
    end;
    cptIsSoundHint: begin
      _CurParameter.IsSoundHint := StrToBool(parameterValue);
    end;
  end;

  if refreshVideo and Assigned(_OnParameterChangeEvent) then
    _OnParameterChangeEvent(_CurParameter, False);
end;

function TfrmCapParameterCfg.GetRealVideoSize: TVideoSize;
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

  if not Assigned(_capSource) then Exit;

  //查找已经连接的PIN
  hr := TDS9Ex.FindConnectedPin(_capSource, PINDIR_OUTPUT, pin);
  if Failed(hr) then Exit;

  try
    hr := pin.QueryInterface(IID_IAMStreamConfig, amStreamConfig);
    if FAILED(hr) then Exit;

    try
      hr := amStreamConfig.GetFormat(pmt);   //取得当前视频格式
      if FAILED(hr) then Exit;

      pvih := pmt.pbFormat;
      curSize.Width := pvih^.bmiHeader.biWidth;
      curSize.Height := pvih^.bmiHeader.biHeight;
      Result := curSize;

      DeleteMediaType(pmt);

    finally
      amStreamConfig := nil;
    end;
  finally
    pin := nil;
  end;
end;


procedure TfrmCapParameterCfg.LoadCaptureDevice;
var
  deviceNames: TStringList;
begin
  deviceNames := TStringList.Create;
  try
    //取得视频采集设备名称
    TDS9Ex.GetDeviceNames(CLSID_VideoInputDeviceCategory, deviceNames);

    //添加设备名称
    cbbCaptureDevice.Items.Clear;
    cbbCaptureDevice.Items.AddStrings(deviceNames);

  finally
    FreeAndNil(deviceNames);
  end;
end;

class procedure TfrmCapParameterCfg.ShowCaptureParameterCfg(
  capGraphBuilder2: ICaptureGraphBuilder2;
  capSource: IBaseFilter;
  const cfgFileName: WideString;
  const capParameter: TCaptureParameter;
  const hideCfgItem: Integer;
  const postionType: TCapParameterPostion;
  const parentHandle: HWND;
  callBack: TOnParameterChangeEvent;
  vfwConfigCall: TOnVfwConfigCallEvent);
var
  frmCapParameterCfg: TfrmCapParameterCfg;
begin
  frmCapParameterCfg := TfrmCapParameterCfg.Create(Application{nil});
  try
    frmCapParameterCfg._ICapGraphBuilder2 := capGraphBuilder2;
    frmCapParameterCfg._capSource := capSource;  //2010-12-10 端口选择测试修改
    
    if parentHandle > 0 then begin
      frmCapParameterCfg.ParentWindow := parentHandle;
    end;

    frmCapParameterCfg.InitParameterCfg(cfgFileName, capParameter);
    frmCapParameterCfg.HideParameterCfgItem(hideCfgItem);

    frmCapParameterCfg.PositionType := postionType;

    frmCapParameterCfg.OnParameterChange := callBack;
    frmCapParameterCfg.OnVfwConfigCall := vfwConfigCall;

    frmCapParameterCfg.ShowModal();
  finally
    FreeAndNil(frmCapParameterCfg);
  end;
end;

procedure TfrmCapParameterCfg.FormCreate(Sender: TObject);
begin
  //窗口置顶
  setwindowpos(self.handle,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE  or  SWP_NOSIZE);

  _ICapGraphBuilder2 := nil;
  _capSource := nil;
  
  _VideoFormats := '';
  
  pgcParameters.ActivePageIndex := 0;
  _IsSaveCapParameter := True;

  _SizeControl := TSizeControl.Create(Self);
  _SizeControl.AllowTab := False;
  _SizeControl.OnResized := ImageCutAreaChangeEvent;
  _SizeControl.OnMoved := ImageCutAreaChangeEvent;
end;

procedure TfrmCapParameterCfg.LoadVideoEncoder;
var
  deviceNames: TStringList;
begin
  deviceNames := TStringList.Create;
  try
    //取得视频压缩编码器名称
    TDS9Ex.GetDeviceNames(CLSID_VideoCompressorCategory, deviceNames);

    //添加视频压缩编码器名称
    cbbVideoEncoder.Items.Clear;
    cbbVideoEncoder.Items.Add(''); 
    cbbVideoEncoder.Items.AddStrings(deviceNames);

  finally
    FreeAndNil(deviceNames);
  end;
end;

procedure TfrmCapParameterCfg.cbbCaptureDeviceChange(Sender: TObject);
begin
  try
    if Trim(cbbCaptureDevice.Text) = '' then Exit;
                                 
    //读取视频质量
    LoadVideoQuality(cbbCaptureDevice.Text, False);

    //读取视频端口
    LoadCrossbar1(cbbCaptureDevice.Text);

    //设置采集设备参数
    SetParameterValue(cptCaptureDeviceName, cbbCaptureDevice.Text, True);

    _VideoFormats := GetVideoFormats(cbbCaptureDevice.Text);
  except
    on e:Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.LoadVideoAnalogFromConfigFile();
var
  analogName: String;
  analogCount: Integer;
  i: Integer;
begin
  //读取制式的数量
  analogCount := _IniFile.ReadInteger(Section_VideoAnalogCfg, 'Count', 0);

  if analogCount <= 0 then begin
    _IniFile.WriteInteger(Section_VideoAnalogCfg, 'Count', Length(SysVideoAnalog));
    for i := 0 to Length(SysVideoAnalog) - 1 do begin
      _IniFile.WriteString(Section_VideoAnalogCfg, IntToStr(i + 1), SysVideoAnalog[i]);
    end;

    analogCount := Length(SysVideoAnalog);
  end;

  //读取视频制式
  for i := 0 to analogCount do begin
    analogName := _IniFile.ReadString(Section_VideoAnalogCfg, IntToStr(i), '');
    if Trim(analogName) = '' then Continue;

    cbbAnalogVideo.Items.Append(analogName);
  end;
end;

procedure TfrmCapParameterCfg.FormDestroy(Sender: TObject);
begin
  if Assigned(_IniFile) then FreeAndNil(_IniFile);
  if Assigned(_SizeControl) then FreeAndNil(_SizeControl);
end;

procedure TfrmCapParameterCfg.LoadVideoQuality(const deviceName: WideString; const isLoadDefault: Boolean);
var
  captureFilter: TFilter;
  hr: HRESULT;
  amVideoProcAmp: IAMVideoProcAmp;

  //取得质量配置相关信息
  procedure GetQualityInf(curAmVideoProcAmp: IAMVideoProcAmp;
                   trackBar: TTrackBar;
                   PropertyTag : TVideoProcAmpProperty);
  var
    curHr: HRESULT;
    iMinValue, iMaxValue, iStep, iCurValue, iDefault: Integer;
    iFlags : TVideoProcAmpFlags;
  begin

    //取得视频质量设置的范围
    curHr := curAmVideoProcAmp.GetRange(PropertyTag, iMinValue, iMaxValue, iStep, iDefault, iFlags);
    if not Succeeded(curHr) then begin
      trackBar.Enabled := False;

      Exit;
    end;
              
    trackBar.Min := iMinValue;
    trackBar.Max := iMaxValue;
    trackBar.Frequency := iStep;
    trackBar.Position := iDefault;
    trackBar.Tag := Integer(iFlags);

    //取得当前值
    curHr := amVideoProcAmp.Get(PropertyTag, iCurValue, iFlags);
    if not Succeeded(curHr) then begin
      trackBar.Enabled := False;
      Exit;
    end;
    
    //判断是否需要载入默认值
    if not isLoadDefault then begin
      trackBar.Position := iCurValue;
    end else begin
      trackBar.Position := iDefault;
    end;

    trackBar.Enabled := True;
  end;

  
begin
  //对于vfw的设备,则不进行读取
  if TDS9Ex.IsVfwDevice(deviceName) then Exit;

  captureFilter := GetDeviceFilter(deviceName);
  if not Assigned(captureFilter) then Exit;

  try
    //查询filter接口，判断是否支持质量设置
    hr := captureFilter.BaseFilter.CreateFilter.QueryInterface(IID_IAMVideoProcAmp, amVideoProcAmp);
    if not Succeeded(hr) then Exit;

    //说明：经测试在directshow中VideoProcAmp_Flags_Auto表示手动管理，  VideoProcAmp_Flags_Manual表示自动，产生此问题是因为值定义错误
    //VideoProcAmp_Flags_Manual + VideoProcAmp_Flags_Auto = 3
    
    //亮度
    GetQualityInf(amVideoProcAmp, trckbrVideoBrightness, VideoProcAmp_Brightness);
    lblBrightnessValue.Caption := IntToStr(trckbrVideoBrightness.Position);
    lblVideoBrightness.Enabled := trckbrVideoBrightness.Enabled;
    chkIsAutoBrightness.Enabled := IfThen(trckbrVideoBrightness.Tag = 3, 1, 0) > 0;

    //对比度
    GetQualityInf(amVideoProcAmp, trckbrVideoContrast, VideoProcAmp_Contrast);
    lblContrast.Caption := IntToStr(trckbrVideoContrast.Position);
    lblVideoContrast.Enabled := trckbrVideoContrast.Enabled;
    chkIsAutoContrast.Enabled := IfThen(trckbrVideoContrast.Tag = 3, 1, 0) > 0;

    //色调
    GetQualityInf(amVideoProcAmp, trckbrVideoHue, VideoProcAmp_Hue);
    lblHue.Caption := IntToStr(trckbrVideoHue.Position);
    lblVideoHue.Enabled := trckbrVideoHue.Enabled;
    chkIsAutoHue.Enabled := IfThen(trckbrVideoHue.Tag = 3, 1, 0) > 0;

    //饱和度
    GetQualityInf(amVideoProcAmp, trckbrVideoSaturation, VideoProcAmp_Saturation);
    lblSaturation.Caption := IntToStr(trckbrVideoSaturation.Position);
    lblVideoSaturation.Enabled := trckbrVideoSaturation.Enabled;
    chkIsAutoSaturation.Enabled := IfThen(trckbrVideoSaturation.Tag = 3, 1, 0) > 0;

    //伽马
    GetQualityInf(amVideoProcAmp, trckbrVideoGamma, VideoProcAmp_Gamma);
    lblGammaValue.Caption := IntToStr(trckbrVideoGamma.Position);
    lblGamma.Enabled := trckbrVideoGamma.Enabled;
    chkIsAutoGamma.Enabled := IfThen(trckbrVideoGamma.Tag = 3, 1, 0) > 0;

    //白平衡
    GetQualityInf(amVideoProcAmp, trckbrVideoWhiteBlance, VideoProcAmp_WhiteBalance);
    lblWhiteBlanceValue.Caption := IntToStr(trckbrVideoWhiteBlance.Position);
    lblWhiteBlance.Enabled := trckbrVideoWhiteBlance.Enabled;
    chkIsAutoWhiteBlance.Enabled := IfThen(trckbrVideoWhiteBlance.Tag = 3, 1, 0) > 0;

    amVideoProcAmp := nil;
  finally
    FreeAndNil(captureFilter);
  end;
end;

procedure TfrmCapParameterCfg.trckbrVideoBrightnessChange(Sender: TObject);
begin
  //亮度
  try
    lblBrightnessValue.Caption := IntToStr(trckbrVideoBrightness.Position);
    SetParameterValue(cptBrightness, IntToStr(trckbrVideoBrightness.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.trckbrVideoContrastChange(Sender: TObject);
begin
  //对比度
  try
    lblContrast.Caption := IntToStr(trckbrVideoContrast.Position);
    SetParameterValue(cptContrast, IntToStr(trckbrVideoContrast.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;

procedure TfrmCapParameterCfg.trckbrVideoHueChange(Sender: TObject);
begin
  //色调
  try
    lblHue.Caption := IntToStr(trckbrVideoHue.Position);
    SetParameterValue(cptHue, IntToStr(trckbrVideoHue.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;

procedure TfrmCapParameterCfg.trckbrVideoSaturationChange(Sender: TObject);
begin
  //饱和度
  try
    lblSaturation.Caption := IntToStr(trckbrVideoSaturation.Position);
    SetParameterValue(cptSaturation, IntToStr(trckbrVideoSaturation.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;

procedure TfrmCapParameterCfg.trckbrVideoGammaChange(Sender: TObject);
begin
  //伽马
  try
    lblGammaValue.Caption := IntToStr(trckbrVideoGamma.Position);
    SetParameterValue(cptGamma, IntToStr(trckbrVideoGamma.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;

procedure TfrmCapParameterCfg.trckbrVideoWhiteBlanceChange(
  Sender: TObject);
begin
  //白平衡
  try
    lblWhiteBlanceValue.Caption := IntToStr(trckbrVideoWhiteBlance.Position);
    SetParameterValue(cptWhiteBlance, IntToStr(trckbrVideoWhiteBlance.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;
             
procedure TfrmCapParameterCfg.btnDefaultClick(Sender: TObject);
begin
  try
    LoadVideoQuality(cbbCaptureDevice.Text, True);

    //设置视频质量
    SetParameterValue(cptBrightness, IntToStr(trckbrVideoBrightness.Position), false);
    SetParameterValue(cptContrast, IntToStr(trckbrVideoContrast.Position), false);
    SetParameterValue(cptHue, IntToStr(trckbrVideoHue.Position), false);
    SetParameterValue(cptSaturation, IntToStr(trckbrVideoSaturation.Position), false);
    SetParameterValue(cptGamma, IntToStr(trckbrVideoGamma.Position),false);
    SetParameterValue(cptWhiteBlance, IntToStr(trckbrVideoWhiteBlance.Position), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;

procedure TfrmCapParameterCfg.btnCancelClick(Sender: TObject);
begin
  try
    //撤销配置
    _IsSaveCapParameter := False;
    
    Self.Close;
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;  
  end;  
end;

procedure TfrmCapParameterCfg.cbbAnalogVideoChange(Sender: TObject);
begin
  //设置视频制式
  if Trim(cbbAnalogVideo.Text) = '' then Exit;

  try
    SetParameterValue(cptVideoAnalog, cbbAnalogVideo.Text, True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.cbbVideoEncoderChange(Sender: TObject);
begin
  //设置编码器名称
  if Trim(cbbVideoEncoder.Text) = '' then Exit;

  try
    SetParameterValue(cptEncoderName, cbbVideoEncoder.Text, true);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;  
end;

procedure TfrmCapParameterCfg.chkTimeLimitClick(Sender: TObject);
begin
  //是否限时
  try
    edtVideoCaptureTimes.Enabled := chkTimeLimit.Checked;
    SetParameterValue(cptIsTimeLimit, BoolToStr(chkTimeLimit.Checked, True), False);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.edtVideoCaptureTimesChange(Sender: TObject);
begin
  //设置限时模式的时间长度
  try
    if StrToInt(edtVideoCaptureTimes.Text) <= 0 then Exit;
    
    SetParameterValue(cptLimitLength, edtVideoCaptureTimes.Text, False);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
end;

procedure TfrmCapParameterCfg.chkConvertToGrayClick(Sender: TObject);
begin
  //是否转换为8位图
  SetParameterValue(cptIsConvert8Bit, BoolToStr(chkConvertToGray.Checked, True), False);
end;

class procedure TfrmCapParameterCfg.ReadCaptureParameterFromFile(
  const filename: WideString; var captureParameter: TCaptureParameter; var otherPar: TOtherPar);
var
  curIniFile: TIniFile;
begin
  try
    if Trim(filename) = '' then begin
      curIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + CaptureCfgFileName);
    end else begin
      curIniFile := TIniFile.Create(filename);
    end;

    try
      //从ini文件中读取参数
      captureParameter.CaptureDeviceName := curIniFile.ReadString(Section_ParameterCfg, 'CaptureDeviceName', '');
      captureParameter.VideoAnalog := curIniFile.ReadString(Section_ParameterCfg, 'VideoAnalog', '');
      captureParameter.ColorDepth := curIniFile.ReadInteger(Section_ParameterCfg, 'ColorDepth', 24);
      captureParameter.VideoSize := curIniFile.ReadString(Section_ParameterCfg, 'VideoSize', '320X240');
      
      captureParameter.Brightness := curIniFile.ReadInteger(Section_ParameterCfg, 'Brightness', -1);
      captureParameter.Contrast := curIniFile.ReadInteger(Section_ParameterCfg, 'Contrast', -1);
      captureParameter.Hue := curIniFile.ReadInteger(Section_ParameterCfg, 'Hue', -1);
      captureParameter.Saturation := curIniFile.ReadInteger(Section_ParameterCfg, 'Saturation', -1);
      captureParameter.Gamma := curIniFile.ReadInteger(Section_ParameterCfg, 'Gamma', -1);
      captureParameter.WhiteBlance := curIniFile.ReadInteger(Section_ParameterCfg, 'WhiteBlance', -1);

      captureParameter.EncoderName := curIniFile.ReadString(Section_ParameterCfg, 'EncoderName', '');
      captureParameter.IsTimeLimit := curIniFile.ReadBool(Section_ParameterCfg, 'IsTimeLimit', False);
      captureParameter.LimitLength := curIniFile.ReadInteger(Section_ParameterCfg, 'LimitLength', 60);
      captureParameter.IsConvertGrayImg := curIniFile.ReadBool(Section_ParameterCfg, 'IsConvert8Bit', False);
      captureParameter.IsApplyImageCut := curIniFile.ReadBool(Section_ParameterCfg, 'IsApplyImageCut', False);

      captureParameter.TopRate := curIniFile.ReadFloat(Section_ParameterCfg, 'TopRate', 0);
      captureParameter.HeightRate := curIniFile.ReadFloat(Section_ParameterCfg, 'HeightRate', 0);
      captureParameter.LeftRate := curIniFile.ReadFloat(Section_ParameterCfg, 'LeftRate', 0);
      captureParameter.WidthRate := curIniFile.ReadFloat(Section_ParameterCfg, 'WidthRate', 0);

      captureParameter.VideoShowModel := curIniFile.ReadInteger(Section_ParameterCfg, 'VideoShowModel', 0);
      captureParameter.SnatchWay := curIniFile.ReadInteger(Section_ParameterCfg, 'SnatchWay', 0);
      captureParameter.IsShowState := curIniFile.ReadBool(Section_ParameterCfg, 'IsShowState', True);
      
      captureParameter.InputCrossbar := curIniFile.ReadInteger(Section_ParameterCfg, 'InputCrossbar', -1);
      captureParameter.OutputCrossbar := curIniFile.ReadInteger(Section_ParameterCfg, 'OutputCrossbar', -1);

      captureParameter.ExposureWay := curIniFile.ReadInteger(Section_ParameterCfg, 'ExposureWay', 0);    //0不设置，1自动，2手动
      captureParameter.ExposureValue := curIniFile.ReadInteger(Section_ParameterCfg, 'ExposureValue', 0);  //曝光时间

      captureParameter.IsAutoBrightness := curIniFile.ReadBool(Section_ParameterCfg, 'IsAutoBrightness', False);
      captureParameter.IsAutoContrast := curIniFile.ReadBool(Section_ParameterCfg, 'IsAutoContrast', False);
      captureParameter.IsAutoHue := curIniFile.ReadBool(Section_ParameterCfg, 'IsAutoHue', False);
      captureParameter.IsAutoGamma := curIniFile.ReadBool(Section_ParameterCfg, 'IsAutoGamma', False);
      captureParameter.IsAutoSaturation := curIniFile.ReadBool(Section_ParameterCfg, 'IsAutoSaturation', False);
      captureParameter.IsAutoWhiteBlance := curIniFile.ReadBool(Section_ParameterCfg, 'IsAutoWhiteBlance', False);

      captureParameter.IsSoundHint := curIniFile.ReadBool(Section_ParameterCfg, 'IsSoundHint', False);

      captureParameter.DebugFilter := curIniFile.ReadBool(Section_ParameterCfg, 'DebugFilter', False);

      otherPar.Frame := curIniFile.ReadInteger(Section_ParameterCfg, 'Frame', 0);

      captureParameter.ParameterState := True;
    finally
      FreeAndNil(curIniFile);
    end;
  except
    on e: Exception do begin
      //保证参数值完整
      TCaptureParameterConfig.InitCaptureParameter(captureParameter);
      raise e;
    end;
  end;
end;

procedure TfrmCapParameterCfg.btnSureClick(Sender: TObject);
begin
  try
    //保持修改的配置参数到文件
    WriteCaptureParameterToFile(_CaptureParameterCfgFileName, _CurParameter);

    //保存配置并退出
    _IsSaveCapParameter := True;

    Self.Close;
  except
    on e: Exception do begin
      Application.MessageBox(Pchar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end
end;

function TfrmCapParameterCfg.GetDeviceFilter(const deviceName: WideString): TFilter;
var
  capEnum: TSysDevEnum;
  captureFilter: TFilter;
  deviceIndex, i: Integer;
begin
  capEnum := TSysDevEnum.Create(CLSID_VideoInputDeviceCategory);
  try
    //取得指定采集设备的设备索引
    deviceIndex := -1;
    for i := 0 to capEnum.CountFilters - 1 do begin
      if UpperCase(capEnum.Filters[i].FriendlyName) = UpperCase(deviceName) then begin
        deviceIndex := i;
        Break;
      end;
    end;

    //没有找到对应的设备
    if deviceIndex < 0 then begin
      Result := nil;
      Exit;
    end;

    //创建设备关联
    captureFilter := TFilter.Create(nil);
    captureFilter.BaseFilter.Moniker := capEnum.GetMoniker(deviceIndex);

    Result := captureFilter;
  finally
    FreeAndNil(capEnum);
  end;
end;

procedure TfrmCapParameterCfg.tsImageCaptureShow(Sender: TObject);
begin
  _OnParameterChangeEvent(_CurParameter, True);

  LoadCuptureImage();
  ZoomImageSize();

  SetImageDefaultAdjustArea();
  LoadImageCutArea();

  _SizeControl.Target := shapArea;
end;

procedure TfrmCapParameterCfg.LoadCuptureImage;
begin
  if not FileExists(GetCaptureSampleFile) then begin
    cbApplyImgCut.Enabled := False;
    Exit;
  end;


  imgCapture.Picture.LoadFromFile(GetCaptureSampleFile);
  cbApplyImgCut.Enabled := True;
end;

procedure TfrmCapParameterCfg.SetImageDefaultAdjustArea;
begin
  shapArea.Left := imgCapture.Left;
  shapArea.Top := imgCapture.Top;
  shapArea.Height := imgCapture.Height;
  shapArea.Width := imgCapture.Width;

  _SizeControl.SetBounds(shapArea.Left, shapArea.Top, shapArea.Width, shapArea.Height); 


  labRightValue.Caption := '100%';
  labDownValue.Caption := '100%';
end;

procedure TfrmCapParameterCfg.cbApplyImgCutClick(Sender: TObject);
begin
  //设置裁剪参数
  SetParameterValue(cptIsApplyImageCut, BoolToStr(cbApplyImgCut.Checked, True), False);
end;

procedure TfrmCapParameterCfg.FormShow(Sender: TObject);
begin
  {对应位置示意图：
      |-----------------------------------------------------|
      | LeftTop             TopCenter            RightTop   |
      |                                                     |
      |                                                     |
      | LeftCenter         ScreenCenter         RightCenter |
      |                                                     |
      |                                                     |
      | LeftBottom         BottomCenter         RightBottom |
      |-----------------------------------------------------|
  }
  
  case _PositionType of
    cppLeftTop: begin
      Left := Screen.WorkAreaLeft;
      Top := Screen.WorkAreaTop;
    end;
    cppTopCenter: begin
      Left := (Screen.WorkAreaWidth - Self.Width) div 2;
      Top := Screen.WorkAreaTop;
    end;
    cppRightTop: begin
      Left := Screen.WorkAreaWidth - Self.Width;
      Top := Screen.WorkAreaTop;
    end;
    cppRightCenter: begin
      Left := Screen.WorkAreaWidth - Self.Width;
      Top := (Screen.WorkAreaHeight - Self.Height) div 2;
    end;
    cppRightBottom: begin
      Left := Screen.WorkAreaWidth - Self.Width;
      Top := Screen.WorkAreaHeight - Self.Height;
    end;
    cppBottomCenter: begin
      Left := (Screen.WorkAreaWidth - Self.Width) div 2;
      Top := Screen.WorkAreaHeight - Self.Height;
    end;
    cppLeftBottom: begin
      Left := Screen.WorkAreaLeft;
      Top := Screen.WorkAreaHeight - Self.Height;
    end;
    cppLeftCenter: begin
      Left := Screen.WorkAreaLeft;
      Top := (Screen.WorkAreaHeight - Self.Height) div 2;
    end;
    cppScreenCenter: begin
      Left := (Screen.WorkAreaWidth - Self.Width) div 2;
      Top := (Screen.WorkAreaHeight - Self.Height) div 2;
    end;
  end; 
  
  if not _CurParameter.ParameterState then Exit;
    
  //刷新采集窗口
  if Assigned(_OnParameterChangeEvent) then begin
    _OnParameterChangeEvent(_CurParameter, False);
  end;
end;

class function TfrmCapParameterCfg.GetCaptureSampleFile: WideString;
const
  sampleFile: String = 'TMP_592AB3E8CB084DCB8351EBE6A8E54985.bmp';
begin
  Result := ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR + sampleFile;

  //如果目录不存在，则创建目录
  if not FileExists(ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR) then
    ForceDirectories(ExtractFilePath(Application.ExeName) + CONST_TEMP_DIR);
end;

procedure TfrmCapParameterCfg.ZoomImageSize;
var
  rate: Double;
begin

  //取得缩放比例
  if imgCapture.Picture.Bitmap.Height > imgCapture.Picture.Bitmap.Width then begin
    rate := AdjustImageSize / imgCapture.Picture.Bitmap.Height;
  end else begin
    rate := AdjustImageSize / imgCapture.Picture.Bitmap.Width;
  end;

  //调整裁剪样品图像的位置
  imgCapture.Height := Round(rate * imgCapture.Picture.Bitmap.Height);
  imgCapture.Width := Round(rate * imgCapture.Picture.Bitmap.Width);

  if imgCapture.Height < AdjustImageSize then
    imgCapture.Top := (panCut.Height - imgCapture.Height) div 2 + 2;

  if imgCapture.Width < AdjustImageSize then
    imgCapture.Left := panCut.Width div 2 + 2;
end;

procedure TfrmCapParameterCfg.Button1Click(Sender: TObject);
begin
  SetImageDefaultAdjustArea();

  //从设裁剪参数
  SetParameterValue(cptTopRate, FloatToStr((shapArea.Top - imgCapture.Top) / imgCapture.Height), False);
  SetParameterValue(cptHeightRate, FloatToStr(shapArea.Height / imgCapture.Height), False);
  SetParameterValue(cptLeftRate, FloatToStr((shapArea.Left - imgCapture.Left) / imgCapture.Width), False);
  SetParameterValue(cptWidthRate, FloatToStr(shapArea.Width / imgCapture.Width), False);  
end;

procedure TfrmCapParameterCfg.ImageCutAreaChangeEvent(Sender: TObject; ControlRect: TRect);
const
  MinWidth: Integer = 10;
  MinHeight: Integer = 10;
begin
  if shapArea.Height > imgCapture.Height then begin
    //控制裁剪范围不能超出原图大小(纵向)
    shapArea.Top := imgCapture.Top;
    shapArea.Height := imgCapture.Height;

    TSizeControl(Sender).Top := shapArea.Top;
    TSizeControl(Sender).Height := shapArea.Height;
  end;

  if shapArea.Width > imgCapture.Width then begin
    //控制裁剪范围不能超出原图大小(横向)
    shapArea.Left := imgCapture.Left;
    shapArea.Width := imgCapture.Width;

    TSizeControl(Sender).Left := shapArea.Left;
    TSizeControl(Sender).Width := shapArea.Width;
  end;

  if shapArea.Left < imgCapture.Left then begin
    shapArea.Left := imgCapture.Left;
    TSizeControl(Sender).Left := imgCapture.Left;
  end;


  if shapArea.Top < imgCapture.Top then begin
    shapArea.Top := imgCapture.Top;
    TSizeControl(Sender).Top := imgCapture.Top;
  end;


  if shapArea.Left + shapArea.Width > imgCapture.Width + imgCapture.Left then begin
    shapArea.Left := imgCapture.Width - shapArea.Width + imgCapture.Left;
    TSizeControl(Sender).Left := shapArea.Left;
  end;

  if shapArea.Top + shapArea.Height > imgCapture.Height + imgCapture.Top then begin
    shapArea.Top := imgCapture.Height - shapArea.Height + imgCapture.Top;
    TSizeControl(Sender).Top := shapArea.Top;
  end;

  if shapArea.Height < MinHeight then begin
    shapArea.Height := MinHeight;
    TSizeControl(Sender).Height := MinHeight;
  end;

  if shapArea.Width < MinWidth then begin
    shapArea.Width := MinWidth;
    TSizeControl(Sender).Width := MinWidth;
  end;  

  labRightValue.Caption := IntToStr(Round(shapArea.Width / imgCapture.Width * 100)) + '%';
  labDownValue.Caption := IntToStr(Round(shapArea.Height / imgCapture.Height * 100)) + '%';
  
  //设置裁剪参数
  SetParameterValue(cptTopRate, FloatToStr((shapArea.Top - imgCapture.Top) / imgCapture.Height), False);
  SetParameterValue(cptHeightRate, FloatToStr(shapArea.Height / imgCapture.Height), False);
  SetParameterValue(cptLeftRate, FloatToStr((shapArea.Left - imgCapture.Left) / imgCapture.Width), False);
  SetParameterValue(cptWidthRate, FloatToStr(shapArea.Width / imgCapture.Width), False);
end;

procedure TfrmCapParameterCfg.LoadImageCutArea;
begin
  cbApplyImgCut.Checked := _CurParameter.IsApplyImageCut;

  if _CurParameter.HeightRate > 0 then begin
    shapArea.Top := imgCapture.Top + Round(_CurParameter.TopRate * imgCapture.Height);
    shapArea.Height := Round(_CurParameter.HeightRate * imgCapture.Height);
  end;

  if _CurParameter.WidthRate > 0 then begin
    shapArea.Left := imgCapture.Left + Round(_CurParameter.LeftRate * imgCapture.Width);
    shapArea.Width := Round(_CurParameter.WidthRate * imgCapture.Width);
  end;  

  _SizeControl.SetBounds(shapArea.Left, shapArea.Top, shapArea.Width, shapArea.Height);

  labRightValue.Caption := IntToStr(Round(shapArea.Width / imgCapture.Width * 100)) + '%';
  labDownValue.Caption := IntToStr(Round(shapArea.Height / imgCapture.Height * 100)) + '%';  
end;

procedure TfrmCapParameterCfg.btnVideoSourcePropertyClick(Sender: TObject);
var
  errMsg: WideString;
begin
  if not Assigned(_OnVfwConfigCallEvent) then Exit;

  case TControl(Sender).Tag of
    0: begin
       _OnVfwConfigCallEvent(vctVideoSourceProperty, Self.Handle, errMsg);

       //重新读取该方式设置的数据值
       //LoadVideoQuality(_CurParameter.CaptureDeviceName, False);
    end;
    1: begin
      _OnVfwConfigCallEvent(vctVideoCapturePinProperty, Self.Handle, errMsg);
    end;  
    2: _OnVfwConfigCallEvent(vctVfwVideoFormat, Self.Handle, errMsg);
    3: _OnVfwConfigCallEvent(vctVfwVideoDisplay, Self.Handle, errMsg);
    4: _OnVfwConfigCallEvent(vctVideoCrossbar, Self.Handle, errMsg);
    5: _OnVfwConfigCallEvent(vctVfwCompressDialog, Self.Handle, errMsg);
  end;

  if errMsg <> '' then begin
    Application.MessageBox(PAnsiChar('尚未设置采集所需参数，是否现在进行设置？[ERR:' + String(errMsg) + ']'), '提示', MB_OK + MB_ICONINFORMATION);
  end;
end;


procedure TfrmCapParameterCfg.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (ssAlt in Shift) then begin
    //ALT + V
    if Key = Ord('V') then begin
      tsVfwConfig.TabVisible := True;
      pgcParameters.TabIndex := tsVfwConfig.TabIndex;
    end;

    //ALT + W
    if Key = Ord('W') then begin
      tsQuality.TabVisible := True;
      pgcParameters.TabIndex := tsQuality.TabIndex;
    end;

  end;
end;

procedure TfrmCapParameterCfg.pgcParametersChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
  if Trim(cbbCaptureDevice.Text) = '' then begin
    AllowChange := False;
    Application.MessageBox('采集设备不允许为空。', '提示', MB_OK + MB_ICONINFORMATION);
  end;
end;

procedure TfrmCapParameterCfg.LoadVideoFormatFromConfigFile;
var
  formatStr: String;
  formatCount: Integer;
  i: Integer;
begin
  //读取制式的数量
  formatCount := _IniFile.ReadInteger(Section_VideoFormatCfg, 'Count', 0);

  if formatCount <= 0 then begin
    _IniFile.WriteInteger(Section_VideoFormatCfg, 'Count', Length(SysVideoSize));
    for i := 0 to Length(SysVideoSize) - 1 do begin
      _IniFile.WriteString(Section_VideoFormatCfg, IntToStr(i + 1), SysVideoSize[i]);
    end;

    formatCount := Length(SysVideoSize);
  end;

  //读取视频制式
  for i := 0 to formatCount do begin
    formatStr := _IniFile.ReadString(Section_VideoFormatCfg, IntToStr(i), '');
    if Trim(formatStr) = '' then Continue;

    cbxVideoSize.Items.Append(formatStr);
  end;
end;

procedure TfrmCapParameterCfg.cbxVideoSizeChange(Sender: TObject);
var
  realVideoSize: TVideoSize;
  cfgVideoSize: TVideoSize;
  formatCount: Integer;
begin
  //刷新视频格式
  if Trim(cbxVideoSize.Text) = '' then Exit;

  try
    //更改新的分辨率时，需要更新视频
    SetParameterValue(cptVideoSize, cbxVideoSize.Text, true);

    //判断是否成功设置分辨率大小
    realVideoSize := GetRealVideoSize;
    cfgVideoSize := TCaptureParameterConfig.ConvertVideoSizeInf(cbxVideoSize.Text);

    if (realVideoSize.Width <> cfgVideoSize.Width) or (realVideoSize.Height <> cfgVideoSize.Height) then begin
      if Application.MessageBox(Pchar('分辨率设置无效，是否允许恢复到合适的分辨率 [' + IntToStr(realVideoSize.Width) + 'X' + IntToStr(realVideoSize.Height ) + ']。'),
        '提示', MB_YESNO + MB_ICONINFORMATION) = ID_NO then Exit;
        
      cbxVideoSize.ItemIndex := cbxVideoSize.Items.IndexOf(TCaptureParameterConfig.GetVideoSizeStr(realVideoSize.Width, realVideoSize.Height));

      //在分辨率列表中，没有找到适合的分辨率大小
      if cbxVideoSize.ItemIndex < 0 then begin
        cbxVideoSize.Items.Append(TCaptureParameterConfig.GetVideoSizeStr(realVideoSize.Width, realVideoSize.Height));
        cbxVideoSize.ItemIndex := cbxVideoSize.Items.Count - 1;

        //读取制式的数量
        formatCount := _IniFile.ReadInteger(Section_VideoFormatCfg, 'Count', 0);
        formatCount := formatCount + 1;

        //将分辨率写入ini文件保存
        _IniFile.WriteInteger(Section_VideoFormatCfg, 'Count', formatCount);
        _IniFile.WriteString(Section_VideoFormatCfg, IntToStr(formatCount), TCaptureParameterConfig.GetVideoSizeStr(realVideoSize.Width, realVideoSize.Height));
      end;

      //保存恢复后的参数，不需要刷新视频
      SetParameterValue(cptVideoSize, cbxVideoSize.Text, false);
    end;
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.rgpColorDepthClick(Sender: TObject);
begin
  //设置颜色深度
  case rgpColorDepth.ItemIndex of
    0: SetParameterValue(cptColorDepth, IntToStr(8), True);
    1: SetParameterValue(cptColorDepth, IntToStr(24), True);
    2: SetParameterValue(cptColorDepth, IntToStr(12), True);
    3: SetParameterValue(cptColorDepth, IntToStr(32), True);
    4: SetParameterValue(cptColorDepth, IntToStr(16), True);
  end;
end;

class procedure TfrmCapParameterCfg.WriteCaptureParameterToFile(
  const filename: WideString; captureParameter: TCaptureParameter);
var
  curIniFile: TIniFile;
begin
  if Trim(filename) = '' then begin
    curIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + CaptureCfgFileName);
  end else begin
    curIniFile := TIniFile.Create(filename);
  end;

  try
    //写入参数到Ini文件中
    curIniFile.WriteString(Section_ParameterCfg,  'CaptureDeviceName', captureParameter.CaptureDeviceName );
    curIniFile.WriteString(Section_ParameterCfg,  'VideoAnalog', captureParameter.VideoAnalog);
    curIniFile.WriteInteger(Section_ParameterCfg, 'ColorDepth', captureParameter.ColorDepth);
    curIniFile.WriteString(Section_ParameterCfg,  'VideoSize', captureParameter.VideoSize);
    
    curIniFile.WriteInteger(Section_ParameterCfg, 'Brightness', captureParameter.Brightness);
    curIniFile.WriteInteger(Section_ParameterCfg, 'Contrast', captureParameter.Contrast);
    curIniFile.WriteInteger(Section_ParameterCfg, 'Hue', captureParameter.Hue);
    curIniFile.WriteInteger(Section_ParameterCfg, 'Saturation', captureParameter.Saturation);
    curIniFile.WriteInteger(Section_ParameterCfg, 'Gamma', captureParameter.Gamma);
    curIniFile.WriteInteger(Section_ParameterCfg, 'WhiteBlance', captureParameter.WhiteBlance);

    curIniFile.WriteString(Section_ParameterCfg,  'EncoderName', captureParameter.EncoderName);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsTimeLimit', captureParameter.IsTimeLimit);
    curIniFile.WriteInteger(Section_ParameterCfg, 'LimitLength', captureParameter.LimitLength);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsConvert8Bit', captureParameter.IsConvertGrayImg);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsApplyImageCut', captureParameter.IsApplyImageCut);
    
    curIniFile.WriteFloat(Section_ParameterCfg,   'TopRate', captureParameter.TopRate);
    curIniFile.WriteFloat(Section_ParameterCfg,   'HeightRate', captureParameter.HeightRate);
    curIniFile.WriteFloat(Section_ParameterCfg,   'LeftRate', captureParameter.LeftRate);
    curIniFile.WriteFloat(Section_ParameterCfg,   'WidthRate', captureParameter.WidthRate);

    curIniFile.WriteInteger(Section_ParameterCfg,   'VideoShowModel', captureParameter.VideoShowModel);
    curIniFile.WriteInteger(Section_ParameterCfg,   'SnatchWay', captureParameter.SnatchWay);
    curIniFile.WriteBool(Section_ParameterCfg,   'IsShowState', captureParameter.IsShowState);
    
    curIniFile.WriteInteger(Section_ParameterCfg,   'InputCrossbar', captureParameter.InputCrossbar);
    curIniFile.WriteInteger(Section_ParameterCfg,   'OutputCrossbar', captureParameter.OutputCrossbar);

    curIniFile.WriteBool(Section_ParameterCfg,    'IsAutoBrightness', captureParameter.IsAutoBrightness);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsAutoContrast', captureParameter.IsAutoContrast);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsAutoHue', captureParameter.IsAutoHue);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsAutoGamma', captureParameter.IsAutoGamma);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsAutoSaturation', captureParameter.IsAutoSaturation);
    curIniFile.WriteBool(Section_ParameterCfg,    'IsAutoWhiteBlance', captureParameter.IsAutoWhiteBlance);

    curIniFile.WriteBool(Section_ParameterCfg,    'IsSoundHint', captureParameter.IsSoundHint);

    curIniFile.WriteBool(Section_ParameterCfg,    'DebugFilter', captureParameter.DebugFilter);

    //if  (captureParameter.ExposureWay = captureParameter.ExposureValue) and (captureParameter.ExposureWay = 0) then begin
    //  curIniFile.WriteInteger(Section_ParameterCfg,   'ExposureWay', 0);
    //  curIniFile.WriteInteger(Section_ParameterCfg,   'ExposureValue', 0);
    //end;
  finally
    FreeAndNil(curIniFile);
  end;
end;

procedure TfrmCapParameterCfg.butVideoEncoderPropertyClick(
  Sender: TObject);
begin
  try
    if Trim(cbbVideoEncoder.Text) = '' then Exit;

    TDS9Ex.ShowEncoderFilterProperty(cbbVideoEncoder.Text, Self.Handle);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);    
    end;
  end;
end;

procedure TfrmCapParameterCfg.InitParameterCfg(
  const cfgFileName: WideString; const capParameter: TCaptureParameter);
begin
  _IsAllowWriteCapturePar := False;
  _CaptureParameterCfgFileName := cfgFileName;

  //创建INI文件读取对象
  if Trim(_CaptureParameterCfgFileName) = '' then begin
    _IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + CaptureCfgFileName);
  end else begin
    _IniFile := TIniFile.Create(_CaptureParameterCfgFileName);
  end;

  TCaptureParameterConfig.CopyParameter(capParameter, _CurParameter);
  TCaptureParameterConfig.CopyParameter(capParameter, _OldParameter);


  //读取采集设备
  cbbCaptureDevice.Clear;
  LoadCaptureDevice();

  //如果采集设备读取失败或者没有采集设备，则所有配置将不可用
  if cbbCaptureDevice.Items.Count <= 0 then Exit;
  cbbCaptureDevice.ItemIndex := 0;

  //读取采集端口
  LoadCrossbar1(cbbCaptureDevice.Text);

  //读取视频质量
  LoadVideoQuality(cbbCaptureDevice.Text, False);

  //读取视频编码器
  try
    cbbVideoEncoder.Clear;
    LoadVideoEncoder();
    if cbbVideoEncoder.Items.Count > 0 then
      cbbVideoEncoder.ItemIndex := 0;
  except
    on e: Exception do begin
      Application.MessageBox(PChar('视频编码器读取错误：' + e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;    
      
  //读取视频制式
  cbbAnalogVideo.Clear;
  LoadVideoAnalogFromConfigFile();
  if cbbAnalogVideo.Items.Count > 0 then
    cbbAnalogVideo.ItemIndex := 0;


  //读取分辨率配置
  cbxVideoSize.Clear;
  LoadVideoFormatFromConfigFile();
  if cbxVideoSize.Items.Count > 0 then
    cbxVideoSize.ItemIndex := 0;

  //从文件中读取配置参数
  try
    ReadCaptureParameterCfgToFace();
  except
    on e: Exception do begin
      Application.MessageBox(PChar('采集参数读取错误：' + e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;  

  _IsAllowWriteCapturePar := True;

  //取得支持的分辨率格式
  _VideoFormats := GetVideoFormats(cbbCaptureDevice.Text);  
end;

procedure TfrmCapParameterCfg.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  //保存参数配置
  if Assigned(_OnParameterChangeEvent) then begin
    if _IsSaveCapParameter then begin
      _OnParameterChangeEvent(_CurParameter, false);
    end else begin
      _OnParameterChangeEvent(_OldParameter, false);
    end;
  end;

  //自动释放窗口所占用的内存
  //该设置适用于使用show方式显示的窗体对象
  //Action := caFree;
end;

procedure TfrmCapParameterCfg.rdoGupSnatchWayClick(Sender: TObject);
begin
  //设置视频抓取模式
  try
    SetParameterValue(cptSnatchWay, IntToStr(rdoGupSnatchWay.ItemIndex), true);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.rdoGupShowModelClick(Sender: TObject);
begin
  //设置视频显示模式
  try
    SetParameterValue(cptVideoShowModel, IntToStr(rdoGupShowModel.ItemIndex), true);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.chkIsShowStateClick(Sender: TObject);
begin
  //设置视频状态显示
  try
    SetParameterValue(cptIsShowState, BoolToStr(chkIsShowState.Checked), true);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.HideParameterCfgItem(
  const hideCfgItem: Integer);
begin
  //视频显示
  if (hideCfgItem and hciVideoDisplay) > 0 then begin
    tsVideoDisplay.TabVisible := False;
  end;

  //图像采集
  if (hideCfgItem and hciImageCapture) > 0 then begin
    tsImageCapture.TabVisible := False;
  end;

  //高级配置
  if (hideCfgItem and hciAdvanceCfg) > 0 then begin
    tsVfwConfig.TabVisible := False;
  end;

  //显示方式
  if (hideCfgItem and hciVideoShowWay) > 0 then begin
    rdoGupShowModel.Enabled := False;
  end;

  //抓取方式
  if (hideCfgItem and hciVideoSnatchWay) > 0 then begin
    rdoGupSnatchWay.Enabled := False;
  end;

  //视频显示状态
  if (hideCfgItem and hciVideoState) > 0 then begin
    chkIsShowState.Enabled := False;
  end;

  //采集设备
  if (hideCfgItem and hciCaptureDevice) > 0 then begin
    tsDevice.TabVisible := False;
  end;

  //视频质量
  if (hideCfgItem and hciVideoQuality) > 0 then begin
    tsQuality.TabVisible := False;
  end;

  //视频编码
  if (hideCfgItem and hciVideoEncoder) > 0 then begin
    tsVideo.TabVisible := False;
  end;      
      
end;




{procedure TfrmCapParameterCfg.LoadCrossbar;
var
  I: integer;
  hr: HRESULT;
  cOutput, cInput: Longint;
  lRelated : Longint;
  lType : TPhysicalConnectorType;
  IBFilter : IBaseFilter;
  iCrossbar: IAMCrossbar;
begin
  try
    cbxInput.Clear;
    cbxOutput.Clear;

    hr := TDS9Ex.CreateFilterByDeviceCategory(AM_KSCATEGORY_CROSSBAR, IID_IAMCrossbar, IBFilter);
    if not Succeeded(hr) then exit;

    iCrossbar := IBFilter as IAMCrossbar;

    if iCrossbar = nil then exit;

    cOutput := -1;
    cInput := -1;

    hr := iCrossbar.get_PinCounts(cOutput, cInput);

    if Succeeded(hr) then
    begin
      for I := 0 to cOutput - 1 do
      begin
        lType := 0;
        iCrossbar.get_CrossbarPinInfo(False, I, lRelated, lType);
        cbxOutput.Items.Add(IntToStr(I) + ' - ' + TDS9Ex.GetCrossbarPinTypeName(lType));
      end;

      for I := 0 to cInput - 1 do
      begin
        iCrossbar.get_CrossbarPinInfo(True, I, lRelated, lType);
        cbxInput.Items.Add(IntToStr(I) + ' - ' + TDS9Ex.GetCrossbarPinTypeName(lType));
      end
    end;
  except
    on e: Exception do
      ShowMessage(e.Message );
  end;
end;}

procedure TfrmCapParameterCfg.butTestClick(Sender: TObject);
var
  filter_Crossbar: IBaseFilter;
  //iCrossBar: IAMCrossbar;
begin
  filter_Crossbar := CreateComObject(clsid_crossbarfilterpropertypage) as ibasefilter;
  ShowFilterPropertyPage(Self.Handle, filter_Crossbar, ppVFWCapSource);
end;

procedure TfrmCapParameterCfg.LoadCrossbar1(const deviceName: WideString);
//var
  //i: integer;
  //captureFilter: TFilter;
  //pinlist: TPinList;
  //pinInf: TPinInfo;
begin
  LoadCrossbar2(deviceName);
  
  Exit;

  {cbxInput.Clear;
  cbxOutput.Clear;

  //对于vfw的设备,则不进行读取
  if TDS9Ex.IsVfwDevice(deviceName) then Exit;

  captureFilter := GetDeviceFilter(deviceName);
  if not Assigned(captureFilter) then Exit;

  try
    //查询filter接口，判断是否支持质量设置
    pinlist := TPinList.Create(captureFilter.BaseFilter.CreateFilter);

    if Assigned(pinlist) then begin
      for I := 0 to pinlist.Count - 1 do begin
        pinInf := pinlist.PinInfo[i];

        if pinInf.dir = PINDIR_INPUT then cbxInput.Items.Add(IntToStr(cbxInput.Items.Count) + '-' + pinInf.achName);
        if pinInf.dir = PINDIR_OUTPUT then cbxOutput.Items.Add(IntToStr(cbxOutput.Items.Count) + '-' + pinInf.achName);
      end;

      if cbxInput.Items.Count > 0 then cbxInput.ItemIndex := 0;
      if cbxOutput.Items.Count > 0 then cbxOutput.ItemIndex := 0;
    end;
  finally
    FreeAndNil(captureFilter);
  end;}
end;

procedure TfrmCapParameterCfg.cbxInputChange(Sender: TObject);
begin
  //设置输入端口
  if cbxInput.Items.Count <= 0 then Exit;

  try
    SetParameterValue(cptInputCrossbar, IntToStr(cbxInput.ItemIndex), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.cbxOutputChange(Sender: TObject);
begin
  //设置输入端口
  if cbxOutput.Items.Count <= 0 then Exit;

  try
    SetParameterValue(cptOutputCrossbar, IntToStr(cbxOutput.ItemIndex), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.LoadCrossbar2(const deviceName: WideString);
var
  captureFilter: TFilter;
  //capSourceFilter: IBaseFilter;
  pFilter: IBaseFilter;
  iAmCrossbarObj1, iAmCrossbarObj2: IAMCrossbar;
  i, iInput, iOutput: Integer;
  lRelated : Longint;
  lType : TPhysicalConnectorType;
  hr: HRESULT;

  procedure SetDefaultCrossbarList();
  begin
    if cbxInput.Items.Count <= 0 then cbxInput.Items.Append('无');
    if cbxOutput.Items.Count <= 0 then cbxOutput.Items.Append('无');

    cbxInput.ItemIndex := 0;
    cbxOutput.ItemIndex := 0;

    cbxOutput.Enabled := False;
    cbxInput.Enabled := False;
  end;
begin
  cbxInput.Clear;
  cbxOutput.Clear;

  //设置无端口信息时的端口下啦列表的显示状态
  SetDefaultCrossbarList();

  //对于vfw的设备,则不进行读取
  if TDS9Ex.IsVfwDevice(deviceName) then Exit;


  //2010-12-10 端口选择测试修改
  
  //captureFilter := GetDeviceFilter(deviceName);
  //if not Assigned(captureFilter) then Exit;
  // hr := TDS9Ex.CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, deviceName, capSourceFilter);
  //if Failed(hr) then Exit;

  try

    //使用该方法可以获取IID_IAMCrossbar 的Filter
    //hr := _ICapGraphBuilder2.FindInterface(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Video, _capSource, IID_IAMCrossbar, iAMCrossbarObj1);
    hr := _ICapGraphBuilder2.FindInterface(@LOOK_UPSTREAM_ONLY, nil, _capSource, IID_IAMCrossbar, iAmCrossbarObj1);
    //hr := _ICapGraphBuilder2.FindInterface(@LOOK_UPSTREAM_ONLY, nil, captureFilter.BaseFilter.CreateFilter, IID_IAMCrossbar, iAmCrossbarObj1);
    //hr := _ICapGraphBuilder2.QueryInterface(IID_IAMCrossbar, iAmCrossbarObj1);
    //if Failed(hr) then begin
     //hr := _ICapGraphBuilder2.FindInterface(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Video, nil, IID_IAMCrossbar, iAmCrossbarObj);
    //end;
               //ShowMessage('Crossbar1');
    if Succeeded(hr) then begin
      hr := iAmCrossbarObj1.QueryInterface(IID_IBaseFilter, pFilter);
      try
        if Succeeded(hr) then begin
          _ICapGraphBuilder2.FindInterface(@LOOK_UPSTREAM_ONLY, nil, pFilter, IID_IAMCrossbar, iAmCrossbarObj2);
        end;
      finally
        pFilter := nil;
      end;

      hr := S_OK;
    end;

    if Failed(hr) then exit;

    try
      hr := iAmCrossbarObj1.get_PinCounts(iOutput, iInput);
      if Failed(hr) then exit;

      cbxOutput.Clear;
      for I := 0 to iOutput - 1 do
      begin
        lType := 0;
        iAmCrossbarObj1.get_CrossbarPinInfo(False, I, lRelated, lType);
        cbxOutput.Items.Add('cbr1 >> ' + IntToStr(I) + ' - ' + TDS9Ex.GetCrossbarPinTypeName(lType));
      end;
      cbxOutput.Enabled := iOutput > 0;

      cbxInput.Clear;
      for I := 0 to iInput - 1 do
      begin
        iAmCrossbarObj1.get_CrossbarPinInfo(True, I, lRelated, lType);
        cbxInput.Items.Add('cbr1 >> ' + IntToStr(I) + ' - ' + TDS9Ex.GetCrossbarPinTypeName(lType));
      end;
      cbxInput.Enabled := iInput > 0;
      
    finally
      iAmCrossbarObj1 := nil;
    end;
  finally
    FreeAndNil(captureFilter);
  end;
end;

function TfrmCapParameterCfg.GetVideoFormats(
  const deviceName: WideString): String;
var
  captureFilter: TFilter;
  VideoMediaTypes: TEnumMediaType;
  pinList: TPinList;
  i: Integer;
begin
  //对于vfw的设备,则不进行读取
  if TDS9Ex.IsVfwDevice(deviceName) then Exit;
  
  captureFilter := GetDeviceFilter(deviceName);
  if not Assigned(captureFilter) then Exit;

  try

    Result := '';

    pinList := TPinList.Create(captureFilter.BaseFilter.CreateFilter);
    VideoMediaTypes := TEnumMediaType.Create;
    try
      VideoMediaTypes.Assign(pinList.First);
      for i := 0 to VideoMediaTypes.Count - 1 do begin
        Result := Result + VideoMediaTypes.MediaDescription[i];
      end; 
    finally
      FreeAndNil(VideoMediaTypes);
      FreeAndNil(pinList);
    end;

  finally
    FreeAndNil(captureFilter);
  end;

end;

procedure TfrmCapParameterCfg.cbxVideoSizeDrawItem(Control: TWinControl;
  Index: Integer; Rect: TRect; State: TOwnerDrawState);
begin
  cbxVideoSize.Canvas.FillRect(Rect);

  if Trim(_VideoFormats) = '' then begin
    cbxVideoSize.Canvas.TextOut(5, Rect.Top, cbxVideoSize.Items[Index]);
    exit;
  end;

  //cbxVideoSize.Canvas.Font.Style := [fsBold];
  if (pos(cbxVideoSize.Items[Index], _VideoFormats) > 0) then
    cbxVideoSize.Canvas.Font.Color := clGreen
  else
    cbxVideoSize.Canvas.Font.Color := clSilver;

  cbxVideoSize.Canvas.TextOut(5, Rect.Top, cbxVideoSize.Items[Index]);
end;

procedure TfrmCapParameterCfg.chkIsAutoBrightnessClick(Sender: TObject);
begin
  //亮度
  try
    SetParameterValue(cptIsAutoBrightness, BoolToStr(chkIsAutoBrightness.Checked), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.chkIsAutoContrastClick(Sender: TObject);
begin
  //对比度
  try
    SetParameterValue(cptIsAutoContrast, BoolToStr(chkIsAutoContrast.Checked), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.chkIsAutoHueClick(Sender: TObject);
begin
  //色调
  try
    SetParameterValue(cptIsAutoHue, BoolToStr(chkIsAutoHue.Checked), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;


procedure TfrmCapParameterCfg.chkIsAutoSaturationClick(Sender: TObject);
begin
  //饱和度
  try
    SetParameterValue(cptIsAutoSaturation, BoolToStr(chkIsAutoSaturation.Checked), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.chkIsAutoGammaClick(Sender: TObject);
begin
  //伽马
  try
    SetParameterValue(cptIsAutoGamma, BoolToStr(chkIsAutoGamma.Checked), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.chkIsAutoWhiteBlanceClick(Sender: TObject);
begin
  //白平衡
  try
    SetParameterValue(cptIsAutoWhiteBlance, BoolToStr(chkIsAutoWhiteBlance.Checked), True);
  except
    on e: Exception do begin
      Application.MessageBox(PChar(e.Message), '提示', MB_OK + MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmCapParameterCfg.cbSoundHintClick(Sender: TObject);
begin
  //是否转换为8位图
  SetParameterValue(cptIsSoundHint, BoolToStr(cbSoundHint.Checked, True), False);
end;

end.
