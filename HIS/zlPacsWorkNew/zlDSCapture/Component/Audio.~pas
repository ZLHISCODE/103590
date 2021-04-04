{*************************************************************************
  Unit:                Audio.pas

  Description:         TAudio component for accessing waveform devices

  Accessed Units:      mmSystem.pas

  Compiler:            Delphi 7 (32 bit)  2010-11-19

  I/O:                 waveform device via Windows multimedia API

  Author:              tjh

 ....Reference  Base Audio 4.0 Components，ACMWaveIn Components， Voice.Communicator 2.5 Components， Lame Source，MP3export.pas ....

**************************************************************************}

Unit Audio;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Forms, Classes, Graphics,
  mmSystem, Controls, ExtCtrls, Dialogs, LameCompress;

type
  TChannels = (acMono, acStereo);
  TBPS = (bps8, bps16);
  TBufferCount = (buf2, buf4, buf6, buf8, buf10, buf12, buf14, buf16);
  
const
  C_BPS8 = 8;
  C_BPS16 = 16;
  C_MONO = 1;
  C_STEREO = 2;
  
  DefaultAudioDeviceID = WAVE_MAPPER;
  ChannelsDefault = C_STEREO;
  BPSDefault = C_BPS16;
  BufferCountDefault = buf8;

  cypl = 44100; //normally 5512 8000 11050 22100 44100 48000;
  cycd = 1024; //1024


type
  TMciAudioState = (masRun, masStop, masPause);

  //TAudioSettings类描述（基础配置）
  TAudioSettings = class(TCustomControl)
  private
    DeviceOpen           : Boolean;
    _SepCtrl             : Boolean;
    _DeviceId            : UINT;
    _WaveFmtSize         : Integer;
    _MciAudioState       : TMciAudioState;
    _BufferCount          : TBufferCount;
    _WaveBufSize          : Word;
    _ErrorMessage         : string;

    _AudioLineColor: TColor;       //线条显示颜色
    _MaxColor: TColor;        //最大值的显示颜色
    _SampleCount: Integer;    //保存绘制的采样点数
    _WordState: Boolean;
    _DrawAreaIndex: Integer;
    _audioData:array of byte;      //录音时的采样数据
    _Title: String;

    _BufCanvas: Graphics.TBitmap; //缓冲，绘制完毕后复制到界面，避免闪烁
    _DrawTimmer: TTimer;
  protected
    FNoSamples           : Word;
    pWaveFmt             : pWaveFormatEx;
    
    procedure SetChannels(Value:TChannels); virtual;
    procedure SetBPS(Value: TBPS); virtual;
    procedure SetSPS(Value:Word); virtual;
    procedure SetNoSamples(Value:Word); virtual;
    procedure SetDeviceId(const Value: UINT); virtual;
    procedure SetLineColor(const value: TColor);
    procedure SetMaxColor(const value: TColor);
    function GetDrawFrequency: Integer;
    procedure SetDrawFrequency(const Value: Integer);
    procedure SetTitle(const Value: String);
    function GetBPS: TBPS;
    function GetChannels: TChannels;
    function GetSPS: Word;
    function GetFormatTag: Integer;
    procedure SetFormatTag(const Value: Integer);

    function RefreshWaveFormat: Boolean;
    procedure FreeMemory;

    procedure DrawAudioLine(const IsStop: Boolean = false);
    procedure DrawTimerEvent(sender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy(); override;

    procedure ShowFormatDialog();
    procedure LoadDeviceSamples();
    
    property AudioState: TMciAudioState read _MciAudioState;
    property ErrorMsg: String read _ErrorMessage;
  published
    property FormatTag: Integer read GetFormatTag write SetFormatTag;
    property BitsPerSample: TBPS read GetBPS write SetBPS;// default BPSDefault;
    property Channels: TChannels read GetChannels write SetChannels;// default ChannelsDefault;
    property SampleRate: Word read GetSPS write SetSPS;// default cypl;
    property NoSamples: Word read FNoSamples write SetNoSamples;// default cycd;
    property SepCtrl: Boolean read _SepCtrl write _SepCtrl;// default false;
    property DeviceId: UINT read _DeviceId write SetDeviceId;// default DefaultAudioDeviceID;
    property AudioLineColor: TColor read _AudioLineColor write SetLineColor;// default clLime;
    property MaxColor: TColor read _MaxColor write SetMaxColor;// default clRed;
    property SampleCount: Integer read _SampleCount write _SampleCount;// default 200;
    property DrawFrequency: Integer read GetDrawFrequency write SetDrawFrequency; //default 500
    property Title: String read _Title write SetTitle;  //default ''
    property BufferCount: TBufferCount read _BufferCount write _BufferCount; //default buf8
  end;


  TVBRQuality = (vbrQHigh, vbrQ1, vbrQ2, vbrQ3, vbrQ4, vbrQ5, vbrQ6, vbrQ7, vbrQ8, vbrQLow);


  //TLameEnc类描述（音频压缩）
  TLameEnc = class(TObject)
  private
    _lameConfig: BE_CONFIG_FORMATLAME;
    _mp3Stream: TFileStream;  //MP3文件流
    _useLame: Boolean;        //是否使用lame压缩
    _PlanRate: Integer;       //文件压缩时的进度比
    _lameSamples: Integer;    //调用InitStream函数返回lame的采样数
    _fileName: String;        //mp3文件名称

    
    function GetCurTime: Integer;
    function GetStreamPostion: Int64;
    function GetStreamSize: Int64;
    function GetTimeLen: Integer;
    procedure SetCurTime(const Value: Integer);
    procedure SetStreamPostion(const Value: Int64);        //如果指定了文件名，则在使用EncodeStream压缩流时，将自动写入_fileName指向的文件流中
  public
    _minBitRate: Integer;
    _maxBitRate: Integer;
    _avgBitRate: Integer;
    _private: Boolean;
    _crc: Boolean;
    _copyrighted: Boolean;
    _original: Boolean;
    _enableVBR: Boolean;
    _vbrQuality: TVBRQuality;
    _disBRS: Boolean;  
    _sampleRate: Integer;
    _channels: Integer;

    _hstream: HBE_STREAM;
    _minOutputBufSize: Cardinal;
    _outBuf: PByte;
    _outBufUsed: Cardinal;
  public
    constructor Create();
    destructor Destroy(); override;

    function InitStream(): Integer;
    function CloseStream(): Integer;

    procedure LoadWavFormat(const wavFile: String);
    procedure SaveAsFile(const fileName: String);

    function EncodeFile(const inputFile, outputFile: String): Boolean;
    function EncodeStream(data: pointer; const nBytes: Cardinal): Integer;

    property PlanRate: Integer read  _PlanRate;
    property LameSamples: Integer read _lameSamples;

    property StreamSize: Int64 read GetStreamSize;
    property CurTime: Integer read GetCurTime write SetCurTime;
    property TimeLen: Integer read GetTimeLen;
    property StreamPostion: Int64 read GetStreamPostion write SetStreamPostion;
  published
    property UseLame: Boolean read _useLame write _useLame;
    property FileName: String read _fileName write _fileName;
  end;




  //TRecorder类描述（声音录制，可进行压缩）
  TRecorder = class(TAudioSettings)
  private
    WaveIn                   : HWAVEIN;
    FSplit                   : Boolean;
    FTrigLevel               : Word;
    FTriggered               : Boolean;
    RecStream                : TFileStream;
    _wavFileName              : WideString;
    _lame                    : TLameEnc;

    _BufferReadIndex: Integer;
    _HeaderBuffer: array of PWaveHdr;

    procedure SetTrigLevel(Value:Word);
    procedure SetSplit(Value:Boolean);
    function GetStreamSize(): Int64;
    procedure SetRecordFile(value: WideString);
    function GetCurTime: Integer;
    procedure SetCurTime(const value: Integer);
    function GetTimeLen: Integer;

    function GetStreamPostion: Int64;
    procedure SetStreamPostion(const Value: Int64);
    function GetRecordInputCount: Integer;
    function GetRecordInputName(index: Integer): String;

    function TestTrigger(StartPtr:pointer; Size:Word):boolean;
    procedure GetError(iErr : Integer; Additional:string);
    function  OpenDevice : boolean;
    function CloseDevice : boolean;
    procedure ConfigRecordToFile(FileName:string; LP,RP:TStream);
    procedure InitHeaderBuffer();
    procedure ReInitHeaderBuffer();
    procedure SetCompressMp3(const Value: Boolean);
    function GetCompressMp3: Boolean;
    function GetMp3CompressRage: Integer;
    procedure SetMp3CompressRate(const Value: Integer);
  protected
    procedure SetChannels(Value:TChannels); override;
    procedure WaveInCallback(var Msg: TMessage); Message MM_WIM_DATA;
    procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy(); override;

    function  StartRecord : boolean;
    function StopRecord : boolean;
    procedure PauseRecord;
    procedure RestartRecord;
                       
    procedure SaveTmpFile(const fileName: String);

    property StreamSize: Int64 read GetStreamSize;
    property CurTime: Integer read GetCurTime write SetCurTime;
    property TimeLen: Integer read GetTimeLen;
    property StreamPostion: Int64 read GetStreamPostion write SetStreamPostion;
    
    property RecordInputCount: Integer read GetRecordInputCount;
    property RecordInputName[index: Integer]: String read GetRecordInputName;
  published
    property SplitChannels: Boolean read FSplit write SetSplit;
    property TrigLevel: Word read FTrigLevel write SetTrigLevel;
    property Triggered: Boolean read FTriggered write FTriggered;
    property RecordFile: WideString read _wavFileName write SetRecordFile;
    property IsCompressMp3: Boolean read GetCompressMp3 write SetCompressMp3;
    property Mp3CompressRate: Integer read GetMp3CompressRage write SetMp3CompressRate;
    
    property Color;
    property FormatTag;
    property BitsPerSample;// default BPSDefault;
    property Channels;// default ChannelsDefault;
    property SampleRate;// default cypl;
    property NoSamples;// default cycd;
    property SepCtrl;// default false;
    property DeviceId;// default DefaultAudioDeviceID;
    property AudioLineColor;// default clLime;
    property MaxColor;// default clRed;
    property SampleCount;// default 200;
    property DrawFrequency; //default 500
    property Title;  //default ''
    property BufferCount; //default buf8
    property ErrorMsg;
  end;


  //TPlayer类描述（声音播放）
  TPlayer = class(TAudioSettings)
  private
    WaveOut                : HWAVEIN;
    FNoOfRepeats           : Word;
    PlayStream             : TFileStream;
    FPlayFile              : boolean;
    _PlayEndBufferCount: Integer;
    
    _BufferWriteIndex: Integer;
    _HeaderBuffer: array of PWaveHdr;

    procedure CloseDevice();
    function  OpenDevice: boolean;
    procedure GetError(iErr : Integer; Additional:string);

    function GetCurTime: Integer;
    procedure SetTime(const value: Integer);
    function GetTimeLen: Integer;
    function GetStreamSize: Int64;
    function GetStreamPosition: Int64;
    procedure SetStreamPosition(const Value: Int64);
    function GetOutputDeviceCount: Integer;
    function GetOutputDeviceName(index: Integer): String;    

    procedure InitHeaderBuffer;
    procedure ReInitHeaderBuffer;
  protected
    procedure WaveOutCallback(var Msg: TMessage); Message MM_WOM_DONE;
    procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy(); override;

    procedure SetVolume(LeftVolume, RightVolume: Word);
    procedure GetVolume(var LeftVolume, RightVolume: Word);

    procedure StopPlay;
    procedure PausePlay;
    procedure RestartPlay;

    function PlayFile(FileName:string; NoOfRepeats:Word):boolean;

    property CurTime: Integer read GetCurTime write SetTime;
    property TimeLen: Integer read GetTimeLen;
    property StreamSize: Int64 read GetStreamSize;
    property StreamPosition: Int64 read GetStreamPosition write SetStreamPosition;
    property OutputDeviceCount: Integer read GetOutputDeviceCount;
    property OutputDeviceName[index: Integer]: String read GetOutputDeviceName;
  published
    property Color;
    property FormatTag;
    property BitsPerSample;// default BPSDefault;
    property Channels;// default ChannelsDefault;
    property SampleRate;// default cypl;
    property NoSamples;// default cycd;
    property SepCtrl;// default false;
    property DeviceId;// default DefaultAudioDeviceID;
    property AudioLineColor;// default clLime;
    property MaxColor;// default clRed;
    property SampleCount;// default 200;
    property DrawFrequency; //default 500
    property Title;  //default ''
    property BufferCount; //default buf8
    property ErrorMsg; 
  end;


  procedure Register;

  function mrealloc(var data; newSize: Cardinal): pointer; assembler;
  procedure rm(var p: pointer; size: Integer);
  function gcd(a, b: Integer): Integer;
  function max(A, B: Integer): Integer;

  
implementation


uses Windows, DateUtils, Math, MSACM, ConvUtils;


var
  fc:TACMFORMATCHOOSEA;


procedure rm(var p: pointer; size: Integer);
begin
  reallocMem(p, size);
end;  

function mrealloc(var data; newSize: Cardinal): pointer; assembler;
asm
	// IN:
	//	EAX = @data
	//	EDX = newSize

	// OUT:
	//	result = EAX = new pointer

	//  newSize := ((newSize + 511) shr 9) shl 9
	add	edx, 511
	shr	edx, 9
	shl 	edx, 9

	//
	or	edx, edx
	// save eax
	mov	ecx, eax
	je	@@bother	// if newSize = 0

	mov	ecx, eax
	mov	ecx, [ecx]	// ECX = pointer(data)
	or	ecx, ecx
	// save eax
	mov	ecx, eax
	je	@@bother	// or data = nil

	// check if realloc is really required

	mov	eax, [eax]	// EAX = pointer(data)

	//
	// get current size allocated
	// NOTE: THE CODE BELOW DEPENDS ON GETMEM.INC IMPLEMENATION!
	//       (WHICH SEEMS NOT TO BE CHANGED SINCE DELPHI 2)
	//

	sub	eax, 4
	mov	eax, [eax]
	and	eax, $7FFFFFFC
	sub	eax, 4

	// check if new size is the same as allocated
	// for some reason Borland implementation did not perform this check
	cmp	eax, edx

	// restore @data
	mov	eax, ecx

	je	@@noBother	// skip reallocMem call

  @@bother:
	// reallocMem(pointer(data), newSize)
	//
	// due to "smart" implementation of reallocMem, it is not possible to call it directly
	//
	push	ecx
	call	rm
	pop	ecx

  @@noBother:
	// result = [ecx]
	mov	eax, [ecx]
end;


function gcd(a, b: Integer): Integer;
var
  p_max, p_min: Integer;
  x: Integer;
begin
  if (a > b) then begin
    p_max := a;
    p_min := b;
  end
  else begin
    p_max := b;
    p_min := a;
  end;
  //
  if (0 < p_min) then
    repeat
      x := p_max mod p_min;
      p_max := p_min;
      p_min := x;
    until (0 = x)
  else
    p_max := 1;
  //
  result := p_max;
end;



function max(A, B: Integer): Integer;
begin
  if (A < B) then
    result := B
  else
    result := A;
end;

function TAudioSettings.RefreshWaveFormat: Boolean;
begin
    pWaveFmt^.cbSize:=0;
    with pWaveFmt^ do begin
      nBlockAlign := (pWaveFmt^.wBitsPerSample * nChannels) shr 3 ;
      nAvgBytesPerSec := nSamplesPerSec * nBlockAlign;

      _WaveBufSize := FNoSamples;
    end;

    Result:=true;
end;

procedure TAudioSettings.FreeMemory;
begin
  if (pWaveFmt = nil) then Exit
  else begin
    FreeMem(pWaveFmt, _WaveFmtSize);
    pWaveFmt:=nil;
  end;
end;

function TRecorder.TestTrigger(StartPtr:pointer; Size:Word):boolean;
var
    i : longint;
    j :boolean;
    k : Word;
begin
  if not(FTriggered) and (Size>0) then begin
    j:=FTriggered;
    i:=Size;
    k:=FTrigLevel;
    if pWaveFmt^.wBitsPerSample = C_BPS8 then begin
asm
    mov eax,StartPtr
    mov ecx,i
    mov edx,0
@trig8:
    mov dl,[eax]
    cmp dx,k
    jge @out8
    add eax,1
    pop ecx
    loop @trig8
    jmp @out88
@out8:
    mov j,1
@out88:
end;
    end else begin
asm
    mov eax,StartPtr
    mov ecx,i
    shr ecx,1
    mov edx,0
@trig16:
    mov dx,[eax]
    cmp dx,k
    jge @out16
    add eax,2
    loop @trig16
    jmp @out16a
@out16:
    mov j,1
@out16a:
end;
    end;
    FTriggered:=j;
  end;

  Result:=FTriggered;
end;

procedure TRecorder.GetError(iErr : Integer; Additional:string);
var
  pError : PChar;
begin
  try
    GetMem(pError,256);
    try
      waveInGetErrorText(iErr,pError,255);
      _ErrorMessage:=StrPas(pError);
    finally
      FreeMem(pError,256);
    end;

    if length(_ErrorMessage)=0 then begin
      _ErrorMessage:=Additional;
    end else begin
      _ErrorMessage:=Additional+' '+ _ErrorMessage;
    end;
  except end;    
end;

function TPlayer.GetCurTime: Integer;
begin
  if not Assigned(PlayStream) then begin
    Result := 0;
    Exit;
  end;

  Result := PlayStream.Position div pWaveFmt^.nAvgBytesPerSec;
end;

procedure TPlayer.SetTime(const value: Integer);
begin
  if not Assigned(PlayStream) then Exit;
  if PlayStream.Position >= PlayStream.Size then Exit;

  PlayStream.Position := Int64(value) * Int64(pWaveFmt^.nAvgBytesPerSec);
end;


function TPlayer.GetTimeLen: Integer;
begin
  if not Assigned(PlayStream) then begin
    Result := 0;
    Exit;
  end;

  Result := Trunc(PlayStream.Size / pWaveFmt^.nAvgBytesPerSec);
end;

function TPlayer.GetStreamSize: Int64;
begin
  if not Assigned(PlayStream) then begin
    Result := 0;
    Exit;
  end;
  
  Result := PlayStream.Size;
end;

procedure TPlayer.GetError(iErr : Integer; Additional:string);
var
  ErrorText : string;
  pError : PChar;
begin
  GetMem(pError,256);
  waveOutGetErrorText(iErr,pError,255);
  ErrorText:=StrPas(pError);
  FreeMem(pError,256);
  if length(ErrorText)=0 then begin
    _ErrorMessage:=Additional;
  end else begin
    _ErrorMessage:=Additional+' '+ErrorText;
  end;  
end;

function TRecorder.GetStreamSize(): Int64;
begin
  if _lame.UseLame then begin
    Result := _lame.GetStreamSize;
    Exit;
  end;

  if not Assigned(RecStream) then begin
    Result := 0;
    exit;
  end;

  Result := RecStream.Size;
end;

function TRecorder.GetCurTime: Integer;
begin
  if _lame.UseLame then begin
    Result := _lame.GetCurTime;
    Exit;
  end;

  if not Assigned( RecStream ) then begin
    Result := 0;
    Exit;
  end;

  Result := Round(RecStream.Position / pWaveFmt^.nAvgBytesPerSec);
end;

procedure TRecorder.SetCurTime(const value: Integer);
begin
  if _lame.UseLame then begin
    _lame.CurTime := value;
    Exit;
  end;

  if not Assigned( RecStream ) then Exit;

  RecStream.Position := Int64(value) * Int64(pWaveFmt^.nAvgBytesPerSec);
end;

function TRecorder.GetTimeLen: Integer;
begin
  if _lame.UseLame then begin
    Result := _lame.GetTimeLen;
    Exit;
  end;
  
  if not Assigned( RecStream ) then begin
    Result := 0;
    Exit;
  end;

  Result := Trunc(RecStream.Size / pWaveFmt^.nAvgBytesPerSec);
end;

procedure TRecorder.SetRecordFile(value: WideString);
begin
  _wavFileName := value;
  //ConfigRecordToFile(value, nil, nil);
end;

function TRecorder.OpenDevice : boolean;
var
  iErr, i : Integer;
begin
  _ErrorMessage := '';
  
  if not(DeviceOpen) then begin
    Result:=false;

    iErr:=waveInOpen(@WaveIn, DeviceId, pWaveFmt, Handle, 0, CALLBACK_WINDOW or WAVE_MAPPED);

    if (iErr<>0) then begin
      GetError(iErr, 'Could not open the input device for recording: ');
      Exit;
    end;



    DeviceOpen:=true;
    _BufferReadIndex := 0;

    InitHeaderBuffer;

    for i:=0 to  (Integer(_BufferCount) + 1) * 2 - 1 do begin
       {prepare the new block}
       iErr := waveInPrepareHeader(WaveIn, _HeaderBuffer[i], sizeof(TWavehdr));
       if (iErr<>0) then begin
           GetError(iErr,'In Prepare error:');
       end;

          {add it to the buffer}
       iErr:=waveInAddBuffer(WaveIn, _HeaderBuffer[i], sizeof(TWaveHdr)); 
       if iErr<>0 then GetError(iErr, 'Add buffer error:');
    end;

    iErr := waveInStart(WaveIn);
    if (iErr<>0) then begin
      _ErrorMessage:= 'Start error';
    end;
  end;
  
  Result:=true;
end;

function TRecorder.CloseDevice : boolean;
var
  iErr: Integer;
begin
  try
    _ErrorMessage := '';

    iErr := waveinreset(wavein);
    if iErr <> 0 then begin
      _ErrorMessage := 'Wave In Reset Error.';
      Result := False;
      Exit;
    end;

    ReInitHeaderBuffer;

    if (waveInClose(WaveIn)<>0) then begin
      _ErrorMessage:='Error closing input device';
      Result := False;
      Exit;
    end;

    WaveIn := 0;
    DeviceOpen:=false;
    Result:=true;
  except
    on e: Exception do begin
      _ErrorMessage := e.Message;
      Result := False;
    end;
  end;
end;

function TRecorder.StartRecord : boolean;
var
  res: Integer;
begin
  _ErrorMessage := '';
  
  if Trim(_wavFileName) = '' then begin
    _ErrorMessage := 'Invalid Use Of RecordFile。';
    Result := False;
    Exit;
  end;

  if _MciAudioState = masRun then begin
    Result := true;
    exit;     //如果已经运行，则退出
  end;

  if _MciAudioState = masPause then begin   //如果处于暂停状态，则继续运行，并退出
    RestartRecord;
    Result := _MciAudioState = masRun;
        
    Exit;
  end;

  if _lame.UseLame then begin
    _lame.FileName := _wavFileName;
    _lame._sampleRate := pWaveFmt^.nSamplesPerSec;
    _lame._channels := 0; //pWaveFmt^.nChannels;  //0使用立体声
    res := _lame.InitStream();
    if res <> BE_ERR_SUCCESSFUL then begin
      _ErrorMessage := 'Compress Config Faild.';
      Result := False;
      Exit;
    end;

    NoSamples := _lame.LameSamples;
  end else begin
    LoadDeviceSamples;  //更具参数配置获取采样大小
    
    ConfigRecordToFile(_wavFileName, nil, nil);
  end;  

  Result:= OpenDevice;
  if not Result then exit;
  
  _MciAudioState := masRun;
  _DrawTimmer.Enabled := True;
end;

function TRecorder.StopRecord : boolean;
var i:longint;
begin
  _ErrorMessage := '';
  
  if _MciAudioState = masStop then begin
    _ErrorMessage := 'Already Stop Record.';
    Result := True;
    Exit;
  end;
       
  Result:=CloseDevice;
  
  if not _lame.UseLame then begin
    i:=RecStream.Size-8;    { size of file  }
    RecStream.Position:=4;
    RecStream.write(i,4);
    i:=i-$24;               { size of data   }
    RecStream.Position:=40;
    RecStream.write(i,4);

    FreeAndNil(RecStream);
  end else begin
    _lame.CloseStream;
  end;  


   _MciAudioState := masStop;
   _DrawTimmer.Enabled := False;

   DrawAudioLine(True);
end;

procedure TRecorder.PauseRecord;
begin
  if _MciAudioState = masRun then begin
    _MciAudioState := masPause;
  end;  
end;

procedure TRecorder.RestartRecord;
begin
  if _MciAudioState = masPause then begin

    _MciAudioState := masRun;
    _DrawTimmer.Enabled := True;    
  end;  
end;

procedure TRecorder.SaveTmpFile(const fileName: String);
var
  p: PChar;
  tmpStream: TFileStream;
  i: Integer;
begin
  if _lame.UseLame then begin
    _lame.SaveAsFile(fileName);
    Exit;
  end;

  if not Assigned(RecStream) then Exit;
  
  tmpStream := TFileStream.Create(fileName, fmCreate or fmShareDenyNone);
  try

    GetMem(p, RecStream.size);
    try
      RecStream.Position := 0;
      RecStream.ReadBuffer(p^, RecStream.Size);
      RecStream.Position := RecStream.Size;

      tmpStream.Position := 0;
      tmpStream.WriteBuffer(p^, RecStream.Size);
    finally
      FreeMem(p);
    end;

    i:=tmpStream.Size-8;    { size of file  }
    tmpStream.Position:=4;
    tmpStream.write(i,4);
    i:=i-$24;               { size of data   }
    tmpStream.Position:=40;
    tmpStream.write(i,4);
  finally
    FreeAndNil(tmpStream);
  end;  
end;

procedure TRecorder.ConfigRecordToFile(FileName:string; LP,RP:TStream);
var
  temp: string;
begin
  if FileName<>'' then begin
    if Assigned(RecStream) then FreeAndNil(RecStream);

    RecStream := TFileStream.Create(FileName, fmCreate  or fmShareDenyNone);
    temp:='RIFF';RecStream.write(temp[1],length(temp));
    temp:=#0#0#0#0;RecStream.write(temp[1],length(temp));     { File size: to be updated }
    temp:='WAVE';RecStream.write(temp[1],length(temp));
    temp:='fmt ';RecStream.write(temp[1],length(temp));
    temp:=#$10#0#0#0;RecStream.write(temp[1],length(temp));   { Fixed }
    temp:=#1#0;RecStream.write(temp[1],length(temp));         { PCM format }
    
    if pWaveFmt^.nChannels = C_MONO then
      temp:=#1#0
    else temp:=#2#0;
      RecStream.write(temp[1],length(temp));

    RecStream.write(pWaveFmt^.nSamplesPerSec,2);
    temp:=#0#0;RecStream.write(temp[1],length(temp));         { SampleRate is given is dWord }
    with pWaveFmt^ do begin
      RecStream.write(nAvgBytesPerSec,4);
      RecStream.write(nBlockAlign,2);
    end;
    RecStream.write(pWaveFmt^.wBitsPerSample,2);
    temp:='data';RecStream.write(temp[1],length(temp));
    temp:=#0#0#0#0;RecStream.write(temp[1],length(temp));    { Data size: to be updated }
  end;
end;

{ Callback routine used for CALLBACK_FUNCTION in waveOutOpen   }

function TPlayer.OpenDevice: boolean;
var
  iErr : Integer;
begin
  if not(DeviceOpen) then begin
    Result:=false;
    iErr:=waveOutOpen(@WaveOut, _DeviceId, pWaveFmt, Handle, 0, CALLBACK_WINDOW or WAVE_MAPPED);

    if (iErr<>0) then begin
      GetError(iErr,'Could not open the output device for playing: ');
      Exit;
    end;

    InitHeaderBuffer;

    DeviceOpen:=true;
  end;
  Result:=true;
end;

procedure TPlayer.CloseDevice();
var
  iErr: Integer;
begin
  _ErrorMessage := '';
  
  if not(DeviceOpen) then begin
    _ErrorMessage:='Player already closed';
    Exit;
  end;

  iErr := waveOutReset(WaveOut);
  if (iErr<>0) then begin
     GetError(iErr,'Wave Out Reset Error: ');
     Exit;
  end;

  ReInitHeaderBuffer;

  iErr:=waveOutClose(WaveOut);
  if (iErr<>0) then begin
     GetError(iErr,'Error closing output device: ');
     Exit;
  end;

  DeviceOpen:=false;
  FPlayFile:=false;
end;

procedure TPlayer.StopPlay;
begin
    if not(DeviceOpen) then Exit;

    CloseDevice();

    //ReInitHeaderBuffer;

    if Assigned(PlayStream) then begin
      FreeAndNil(PlayStream);
    end;

    _MciAudioState := masStop;
    _DrawTimmer.Enabled := False;

    DrawAudioLine(True);
end;

procedure TPlayer.PausePlay;
begin
  if _MciAudioState <> masRun then exit;

  if DeviceOpen then waveOutPause(WaveOut);
  _MciAudioState := masPause;
end;

procedure TPlayer.RestartPlay;
begin
  if _MciAudioState <> masPause then exit;


  if DeviceOpen then waveOutRestart(WaveOut);
  _MciAudioState := masRun;
  _DrawTimmer.Enabled := True; 
end;

function TPlayer.PlayFile(FileName:string; NoOfRepeats:Word):boolean;
var
  temp:array[0..255] of byte;
  iErr, i : integer;
  Data:word;
  DataSize:longint;
  tmpFile: String;
  PlayFileStream: TFileStream;
begin
  _ErrorMessage := '';
  Result:=false;

  if Trim(FileName) = '' then begin
    _ErrorMessage := 'FileName Is Invalid.';
    Exit;
  end;   

  if _MciAudioState = masPause then begin
    RestartPlay;
    Exit;
  end;

  if _MciAudioState <> masStop then begin
    StopPlay;
  end;

  PlayFileStream:=TFileStream.Create(FileName,fmOpenRead);
  try
    PlayFileStream.Read(temp,22);
    PlayFileStream.Read(temp,2);
    
    if (temp[0]=2) then begin
      if (pWaveFmt^.nChannels <> C_STEREO) then begin
        while FPlayFile do Application.ProcessMessages;
        SetChannels(acStereo);
      end;
    end else begin
      if (pWaveFmt^.nChannels <> C_Mono) then begin
        while FPlayFile do Application.ProcessMessages;
        SetChannels(acMono);
      end;
    end;
    
    PlayFileStream.Read(temp,2);
    Data:=temp[1]*256+temp[0];
    
    if (pWaveFmt^.nSamplesPerSec <> Data) then begin
      while FPlayFile do Application.ProcessMessages;
      SetSPS(Data);
    end;
    
    PlayFileStream.Read(temp,8);
    PlayFileStream.Read(temp,2);
    
    if (temp[0]>8) then begin
      if (pWaveFmt^.wBitsPerSample <> C_BPS16) then begin
        while FPlayFile do Application.ProcessMessages;
        SetBPS(bps16);
      end;
    end else begin
      if (pWaveFmt^.wBitsPerSample <> C_BPS8) then begin
        while FPlayFile do Application.ProcessMessages;
        SetBPS(bps8);
      end;
    end;
    
    PlayFileStream.Read(temp,4); i:=0;
    
    while ((temp[i]<>$64) or (temp[i+1]<>$61) or (temp[i+2]<>$74) or (temp[i+3]<>$61)) do begin
      PlayFileStream.Read(temp[i+4],1);
      inc(i);
    end;
    
    PlayFileStream.Read(DataSize,4);
    FPlayFile:=true;

    if Assigned(PlayStream) then FreeAndNil(PlayStream);


    if OpenDevice then begin
      tmpFile := ExtractFilePath(Application.ExeName) + 'TmpPlay.TMP';

      PlayStream := TFileStream.Create(tmpFile, fmCreate);
    end else begin
      Exit;
    end;

    FNoOfRepeats := NoOfRepeats;
    _BufferWriteIndex := -1;
    _PlayEndBufferCount := 0;
    
    PlayStream.CopyFrom(PlayFileStream, DataSize);
    PlayStream.Position := 0;

    for i:=0 to (Integer(_BufferCount) + 1) * 2 - 1 do begin
      DataSize := PlayStream.Read(_HeaderBuffer[i]^.lpData^, _WaveBufSize{pWaveFmt^.nAvgBytesPerSec});
      _PlayEndBufferCount := i + 1;
      
      if DataSize <= 0 then begin
        //如果不能读取数据，则可能已经播放结束
        Break;
      end;  

      _HeaderBuffer[i]^.dwBufferLength:=DataSize;
      _HeaderBuffer[i]^.dwFlags:=0;
      _HeaderBuffer[i]^.dwLoops:=0;

      iErr:=waveOutPrepareHeader(WaveOut,_HeaderBuffer[i],sizeof(TWAVEHDR));
      if iErr<>0 then begin
        GetError(iErr,'');
        Exit;
      end;

      iErr:=waveOutWrite(WaveOut, _HeaderBuffer[i], sizeof(TWAVEHDR));
      if iErr<>0 then begin
        GetError(iErr,'');
        Exit;
      end;
    end;

    _BufferWriteIndex := 0;
    _MciAudioState := masRun;
    _DrawTimmer.Enabled := True;

    Result:=true;
  finally
    FreeAndNil(PlayFileStream);
  end;
end;

procedure TAudioSettings.SetChannels(Value:TChannels);
begin
  if pWaveFmt^.nChannels <> Integer(Value) + 1 then begin
    pWaveFmt^.nChannels := Integer(Value) + 1;

    RefreshWaveFormat;
  end;
end;

procedure TAudioSettings.SetBPS(Value: TBPS);
begin
  if pWaveFmt^.wFormatTag <> (Integer(Value) + 1) * 8 then begin
    pWaveFmt^.wBitsPerSample := (Integer(Value) + 1) * 8;

    RefreshWaveFormat;
  end;
end;

procedure TAudioSettings.SetSPS(Value:Word);
begin
  if pWaveFmt^.nSamplesPerSec <> Value then begin
    pWaveFmt^.nSamplesPerSec := Value;

    RefreshWaveFormat;
  end;
end;


procedure TRecorder.SetSplit(Value:Boolean);
begin
  if pWaveFmt^.nChannels = C_STEREO then begin
    if FSplit <> Value then FSplit:=Value;
  end else FSplit:=false;
end;

procedure TRecorder.SetTrigLevel(Value:Word);
begin
  if FTrigLevel<>Value then FTrigLevel:=Value;
end;

procedure TPlayer.GetVolume(var LeftVolume,RightVolume:Word);
var
  iErr : Integer;
  Vol : LongInt;
begin
  iErr:=waveOutGetVolume(_DeviceId, @Vol);
  if (iErr<>0) then GetError(iErr,'Get Volume Error:');
  
  LeftVolume:=Word(Vol and $FFFF);
  RightVolume:=Word(Vol shr 16);
end;

procedure TPlayer.SetVolume(LeftVolume, RightVolume:Word);
var
  iErr : Integer;
  Vol : longint;
begin
  Vol:=RightVolume;
  Vol:=(Vol shl 16)+LeftVolume;
  iErr:=waveOutSetVolume(_DeviceID, Vol);
  if (iErr<>0) then GetError(iErr, 'Set Volume Error:');
end;

constructor TRecorder.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DeviceId := 0; //使用CALLBACK_WINDOW or WAVE_MAPPED这种方式打开声音输入设备时，需要指定一个输入设备

  FTrigLevel := 128;
  FTriggered := False;
  FSplit := False;

  _lame := TLameEnc.Create;
  _lame.UseLame := True;  //
end;

destructor TRecorder.Destroy;
begin
  CloseDevice;

  if _lame._hstream <> $FFFFFFFF then _lame.CloseStream;

  FreeAndNil(_lame);
  inherited;
end;

constructor TPlayer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  DeviceId := 0;

  FPlayFile := false;
  PlayStream := nil;
  _BufferWriteIndex := 0;
  _PlayEndBufferCount := 0;
  _Title := '播放'; 
end;

destructor TPlayer.Destroy;
begin
  if Assigned(PlayStream) then FreeAndNil(PlayStream);
  
  inherited;
end;

constructor TAudioSettings.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  acmMetrics(0, ACM_METRIC_MAX_SIZE_FORMAT, _WaveFmtSize);
  _WaveFmtSize := SIZEOF(TWaveFormatEx);
  GetMem(pWaveFmt, _WaveFmtSize);

  pWaveFmt^.wFormatTag:= WAVE_FORMAT_PCM; 
  pWaveFmt^.wBitsPerSample := BPSDefault;   //采样位数 8位，16位
  pWaveFmt^.nChannels := ChannelsDefault;  //通道 单声道：acMono, 立体音：acStereo
  pWaveFmt^.nSamplesPerSec := cypl;       //采样率

  FNoSamples := cycd;

  _SepCtrl := False;
  _DeviceId := DefaultAudioDeviceID;
  _BufferCount := BufferCountDefault;

  Color := clBlack;
  Font.Color := clLime;
  
  _AudioLineColor := clLime;
  _MaxColor := clRed;

  _Title := '';
  _WordState := false;
  _SampleCount := 200;
  _MciAudioState := masStop;

  _DrawAreaIndex := 0;
  _BufCanvas := Graphics.TBitmap.Create;

  _DrawTimmer := TTimer.Create(nil);
  _DrawTimmer.Enabled := False;
  _DrawTimmer.Interval := 500;
  _DrawTimmer.OnTimer := DrawTimerEvent;  

  Width := 250;
  Height := 100;  

  //FreeMemory;
  RefreshWaveFormat;
end;

destructor TAudioSettings.Destroy;
begin
  _MciAudioState := masStop;
  _DrawTimmer.Enabled := False;
  
  FreeAndNil(_DrawTimmer);
  FreeAndNil(_BufCanvas);

  FreeMemory;
  
  inherited;
end;

procedure TRecorder.SetChannels(Value: TChannels);
begin
  inherited;

  SetSplit(FSplit);
end;

procedure TAudioSettings.SetDeviceId(const Value: UINT);
begin
  if _DeviceID<>Value then begin
    if Value>9 then begin
      _DeviceID:=WAVE_MAPPER;
    end else begin
      _DeviceID:=Value;
    end;

    //FreeMemory;
    RefreshWaveFormat;
  end;
end;

procedure TAudioSettings.SetNoSamples(Value: Word);
begin
   if FNoSamples<>Value then begin
    FNoSamples:=Value;
    //FreeMemory;
    RefreshWaveFormat;
  end;
end;


function TRecorder.GetStreamPostion: Int64;
begin
  if _lame.UseLame then begin
    Result := _lame.StreamPostion;
    Exit;
  end;

  if not Assigned(RecStream) then begin
    Result := 0;
    Exit;
  end;

  Result := RecStream.Position;
end;

procedure TRecorder.SetStreamPostion(const Value: Int64);
begin
  if _lame.UseLame then begin
    _lame.StreamPostion := Value;
    Exit;
  end;

  if not Assigned(RecStream) then begin
    Exit;
  end;

  RecStream.Position := Value;
end;

procedure TRecorder.WaveInCallback(var Msg: TMessage);
var
   Header: PWaveHdr;
   iErr, bytesRecorded: integer;
begin
  if not DeviceOpen then Exit;
  try
    if _MciAudioState <> masRun then begin
      //当执行暂停后，不需要对缓存中的数据进行处理，但需要准备数据缓存

      {add it to the buffer}
      iErr:=waveInAddBuffer(WaveIn, _HeaderBuffer[_BufferReadIndex], sizeof(TWaveHdr));
      if iErr<>0 then _ErrorMessage := 'Add buffer error';

      _BufferReadIndex := (_BufferReadIndex + 1) mod ((Integer(_BufferCount) + 1) * 2);

      Exit;
    end;  

    Header:=PWaveHdr(msg.lparam);
    bytesRecorded := header^.dwBytesRecorded;

    if (bytesRecorded > 0) and TestTrigger(header^.lpdata, bytesRecorded) then begin

      if _lame.UseLame then begin
        //实时压缩音频数据
        _lame.EncodeStream(header^.lpdata, bytesRecorded);
      end else begin
        //不压缩音频数据
        RecStream.write(header.lpdata^, bytesRecorded);
      end;    

      try
        setlength(_audioData, bytesRecorded);
        move(header^.lpdata^, _audioData[0], bytesRecorded);
      except
        _audioData := nil;
      end;
      
    end;

    {add it to the buffer}
    iErr:=waveInAddBuffer(WaveIn, _HeaderBuffer[_BufferReadIndex], sizeof(TWaveHdr));
    if iErr<>0 then _ErrorMessage := 'Add buffer error';

    _BufferReadIndex := (_BufferReadIndex + 1) mod ((Integer(_BufferCount) + 1) * 2);
  except end;
end;

procedure TAudioSettings.DrawAudioLine(const IsStop: Boolean);
var
  baseLineH: Integer;
  baseLineW: Integer;
  TextHeight, TextWidth: Integer;


  curDrawDot: array of TPoint;
  curDrawMax, curDrawMin, curDrawMaxPos: Integer;

  procedure GetDrawDot(const drawWay: Integer);
  var
    i, k, curSampleCount: Integer;
    tmp: Integer;
    zoom: Integer;
  begin
      if Length(_audioData) <= 0 then exit;

        SetLength(curDrawDot, _SampleCount + 1);

        curDrawDot[0].X := 0;
        curDrawDot[0].Y := baseLineH;
        curDrawDot[1].X := baseLineW - (_SampleCount div 2);
        curDrawDot[1].Y := baseLineH;

        i := 2;
        curDrawMax := -10000;
        curDrawMin := 10000;

        curDrawMaxPos := baseLineW;
        curSampleCount := Length(_audioData);

        zoom := 1000 div _DrawTimmer.Interval;    //计算一秒中之内需要绘制的次数
        if zoom <= 0 then zoom := 1;

        //计算采样数据是否可以平均划分成zoom个区域
        while Zoom > 1 do begin
          if (Length(_audioData) > _SampleCount * zoom) then Break;
          zoom := zoom - 1;
        end;

        if (Length(_audioData) > _SampleCount * zoom) and (_MciAudioState = masRun) then begin
          if _DrawAreaIndex >= zoom then _DrawAreaIndex := 0;

          k := Round(curSampleCount / zoom) * _DrawAreaIndex;
          zoom := Round(curSampleCount / zoom / _SampleCount);

          _DrawAreaIndex := _DrawAreaIndex + 1;
        end else begin
          k := 0;
          zoom := Round(curSampleCount / _SampleCount);
        end;

        if zoom <= 0 then zoom := 1;

        while k <= length(_audioData) - 1 do begin
          if (_MciAudioState = masStop) then exit;
          tmp := (_audioData[k] - 125);
        
          if curDrawMax < tmp then begin
            curDrawMax := tmp;
            curDrawMaxPos := i;
          end;
          
          if curDrawMin > tmp then curDrawMin := tmp;

          if k mod zoom <> 0 then begin
            k := k + 1;
            Continue;
          end;
        
          curDrawDot[i].y := baseLineH  - tmp;
          curDrawDot[i].x := baseLineW - (_SampleCount div 2 - i);

          k := k + zoom;
          i := i + 1;

          if i + 1 >= _SampleCount then break;
        end;

        curDrawDot[i].X := curDrawDot[i - 1].X + 3;
        curDrawDot[i].Y := baseLineH;
        curDrawDot[i + 1].X := _BufCanvas.Canvas.ClipRect.Right;
        curDrawDot[i + 1].Y := baseLineH;

        SetLength(curDrawDot, i + 2);
  end;

  procedure DrawDot();
  begin
    _BufCanvas.Canvas.Polyline(curDrawDot);

    _BufCanvas.Canvas.Pen.Color := _MaxColor;
    _BufCanvas.Canvas.MoveTo(baseLineW - _SampleCount div 2 + curDrawMaxPos - 4, baseLineH - abs(curDrawMax) - 3);
    _BufCanvas.Canvas.LineTo(baseLineW - _SampleCount div 2 + curDrawMaxPos + 4, baseLineH - abs(curDrawMax) - 3);
    _BufCanvas.Canvas.MoveTo(baseLineW - (_SampleCount div 2 - curDrawMaxPos) - 4, baseLineH - abs(curDrawMax) - 4);
    _BufCanvas.Canvas.LineTo(baseLineW - (_SampleCount div 2 - curDrawMaxPos) + 4, baseLineH - abs(curDrawMax) - 4);
  end;
  
begin
  try
    if (Canvas.ClipRect.Right = 0) or (Canvas.ClipRect.Bottom = 0) then exit;

    _BufCanvas.Width := Width ;
    _BufCanvas.Height := Height;

    _BufCanvas.Canvas.Brush.Style := bsSolid;
    _BufCanvas.Canvas.Brush.Color := Color;
    _BufCanvas.Canvas.Font.Color := _AudioLineColor;
    _BufCanvas.Canvas.pen.Color:=_AudioLineColor;

    _BufCanvas.Canvas.FillRect(_BufCanvas.Canvas.ClipRect);

    baseLineH := _BufCanvas.Canvas.ClipRect.Bottom div 2;
    baseLineW := _BufCanvas.Canvas.ClipRect.Right div 2;

    if IsStop then begin
      TextHeight := _BufCanvas.Canvas.TextHeight('ZLSOFT');
      TextWidth := _BufCanvas.Canvas.TextWidth('ZLSOFT');

      _BufCanvas.Canvas.TextOut(baseLineW - TextWidth div 2, baseLineH - TextHeight - 3, 'ZLSOFT');

      _BufCanvas.Canvas.Font.Color := _AudioLineColor;
      _BufCanvas.Canvas.MoveTo(0, baseLineH);
      _BufCanvas.Canvas.LineTo(_BufCanvas.Canvas.ClipRect.Right, baseLineH);

      //绘制到界面上显示
      Canvas.CopyRect(_BufCanvas.Canvas.ClipRect, _BufCanvas.Canvas, _BufCanvas.Canvas.ClipRect);
      
      exit;
    end;


    //Draw Rec............................
    if (Length(_audioData) > 0) and (_MciAudioState <> masStop) then begin
      GetDrawDot(0);
      
      try
        if Length(curDrawDot) > 0 then DrawDot;
      finally
        curDrawDot := nil;
      end;
    end;

    if _MciAudioState <> masStop then begin
      _BufCanvas.Canvas.TextOut(3, 2, _Title);
      if not _WordState then begin
        _BufCanvas.Canvas.TextOut(3, 2, _Title + ' ●');
      end;

      //if _AudioState = adsPause then _BufCanvas.Canvas.TextOut(3, 2, _Title + ' ●');
    end;

    _WordState := not _WordState;

    //绘制到界面上显示
    Canvas.CopyRect(_BufCanvas.Canvas.ClipRect, _BufCanvas.Canvas, _BufCanvas.Canvas.ClipRect);
  except end;
end;

procedure TAudioSettings.DrawTimerEvent(sender: TObject);
begin
  try
    _DrawTimmer.Enabled := False;

    DrawAudioLine((_MciAudioState = masStop));
    Application.ProcessMessages;

    _DrawTimmer.Enabled := _MciAudioState <> masStop;
  except
    on e: Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

function TAudioSettings.GetDrawFrequency: Integer;
begin
  Result := _DrawTimmer.Interval;
end;

procedure TAudioSettings.SetDrawFrequency(const Value: Integer);
begin
  if Value < 100 then begin
    _DrawTimmer.Interval := 100;
    Exit;
  end;

  _DrawTimmer.Interval := Value;
end;

procedure TAudioSettings.SetLineColor(const value: TColor);
begin
  _AudioLineColor := value;
  
  if not _DrawTimmer.Enabled then DrawAudioLine(_MciAudioState = masStop);
end;

procedure TAudioSettings.SetMaxColor(const value: TColor);
begin
  _MaxColor := value;

  if not _DrawTimmer.Enabled then DrawAudioLine(_MciAudioState = masStop);
end;

procedure TRecorder.WMPaint(var Message: TWMPaint);
begin
  try
    if not _DrawTimmer.Enabled then begin
      DrawAudioLine(_MciAudioState = masStop);
    end;
  except end;

  inherited;
end;

function TRecorder.GetRecordInputCount: Integer;
begin
  Result := waveInGetNumDevs;
end;


function TRecorder.GetRecordInputName(index: Integer): String;
var
  InCaps: PWaveInCaps;
  iErr: Integer;
begin
  InCaps := new(PWaveInCaps);
  try
    iErr := waveInGetDevCaps(index, InCaps, SizeOf(TWaveInCapsA));
    if iErr <> 0 then begin
      Result := '';
      Exit;
    end;

    Result := InCaps^.szPname;

  finally
    Dispose(InCaps);
  end;
end;

procedure TRecorder.InitHeaderBuffer;
var
  i: Integer;
  memBlock: PChar;
  sizeBuf: Integer;
begin
  sizeBuf := _WaveBufSize; //pWaveFmt^.nAvgBytesPerSec;   

  SetLength(_HeaderBuffer, (Integer(_BufferCount) + 1) * 2);
  
  for i:=0 to  (Integer(_BufferCount) + 1) * 2 - 1 do begin
    _HeaderBuffer[i] := New(PWaveHdr);
    GetMem(memBlock, sizebuf); //allocate memory

    _HeaderBuffer[i].lpdata := memBlock;
    _HeaderBuffer[i].dwbufferlength := sizebuf;
    _HeaderBuffer[i].dwbytesrecorded := 0;
    _HeaderBuffer[i].dwUser := 0;
    _HeaderBuffer[i].dwflags := 0;
    _HeaderBuffer[i].dwloops := 0;
  end;
end;

procedure TRecorder.ReInitHeaderBuffer;
var
  i, iErr: Integer;
begin
  for i := 0 to (Integer(_BufferCount) + 1) * 2 - 1 do begin
     iErr := waveInUnprepareHeader(WaveIn, _HeaderBuffer[i], sizeof(TWAVEHDR));
     if (iErr<>0) then begin
       GetError(iErr,'Error in waveInUnprepareHeader:');
     end;

     if Assigned(_HeaderBuffer[i]^.lpData) then FreeMem(_HeaderBuffer[i]^.lpData, _HeaderBuffer[i]^.dwBufferLength);
     _HeaderBuffer[i].lpData := nil;

     if (_HeaderBuffer[i]<>nil) then Dispose(_HeaderBuffer[i]);
    _HeaderBuffer[i]:=nil;
  end;

  _HeaderBuffer := nil;
end;


procedure TPlayer.InitHeaderBuffer;
var
  i: Integer;
  memBlock: PChar;
  sizeBuf: Integer;
begin
  sizeBuf := _WaveBufSize; //}pWaveFmt^.nAvgBytesPerSec;

  SetLength(_HeaderBuffer, (Integer(_BufferCount) + 1 ) * 2);
  
  for i:=0 to  (Integer(_BufferCount) + 1) * 2 - 1 do begin
    _HeaderBuffer[i] := New(PWaveHdr);
    GetMem(memBlock, sizebuf); //allocate memory

    _HeaderBuffer[i].lpdata := memBlock;
    _HeaderBuffer[i].dwbufferlength := sizebuf;
    _HeaderBuffer[i].dwbytesrecorded := 0;
    _HeaderBuffer[i].dwUser := 0;
    _HeaderBuffer[i].dwflags := 0;
    _HeaderBuffer[i].dwloops := 0;
  end;
end;

procedure TPlayer.ReInitHeaderBuffer;
var
  i, iErr: Integer;
begin
  for i := 0 to (Integer(_BufferCount) + 1) * 2 - 1 do begin
     iErr := waveOutUnprepareHeader(WaveOut, _HeaderBuffer[i], sizeof(TWAVEHDR));
     if (iErr<>0) then begin
       GetError(iErr,'Error in waveInUnprepareHeader:');
     end;

     if Assigned(_HeaderBuffer[i]^.lpData) then FreeMem(_HeaderBuffer[i]^.lpData, _HeaderBuffer[i]^.dwBufferLength);
     _HeaderBuffer[i].lpData := nil;

     if (_HeaderBuffer[i]<>nil) then Dispose(_HeaderBuffer[i]);
    _HeaderBuffer[i]:=nil;
  end;

  _HeaderBuffer := nil;
end;

procedure TPlayer.WaveOutCallback(var Msg: TMessage);
var
  bytesRecorded: integer;
  wSize: Int64;
  iErr: integer;
begin
  if not DeviceOpen then Exit;
  if _BufferWriteIndex < 0 then Exit;
  if not Assigned(PlayStream) then Exit; 

  try
    if PlayStream.Position >= PlayStream.Size then begin
      _PlayEndBufferCount := _PlayEndBufferCount - 1;

      if _PlayEndBufferCount <= 0 then begin
        StopPlay;
      end;

      Exit;
    end;

    //读取采样数据,以显示波形
    bytesRecorded := _HeaderBuffer[_BufferWriteIndex]^.dwBufferLength;

    if bytesRecorded > 0 then begin
      try
        setlength(_audioData, bytesRecorded);
        move(_HeaderBuffer[_BufferWriteIndex]^.lpData^, _audioData[0], bytesRecorded);
      except
        _audioData := nil;
      end;
    end;
    
    iErr := waveOutUnprepareHeader(WaveOut, _HeaderBuffer[_BufferWriteIndex], sizeof(TWAVEHDR));
    if iErr <> 0 then begin
      GetError(iErr,'');
    end;

    wSize := PlayStream.Read(_HeaderBuffer[_BufferWriteIndex]^.lpData^, _WaveBufSize{pWaveFmt^.nAvgBytesPerSec});
    if wSize <= 0 then Exit;

    _HeaderBuffer[_BufferWriteIndex]^.dwBufferLength:=wSize;
    _HeaderBuffer[_BufferWriteIndex]^.dwFlags:=0;
    _HeaderBuffer[_BufferWriteIndex]^.dwLoops:=0;

    iErr:=waveOutPrepareHeader(WaveOut,_HeaderBuffer[_BufferWriteIndex], sizeof(TWAVEHDR));
    if iErr<>0 then begin
      GetError(iErr, '');
      Exit;
    end;

    iErr:=waveOutWrite(WaveOut, _HeaderBuffer[_BufferWriteIndex], sizeof(TWAVEHDR));
    if iErr<>0 then begin
      GetError(iErr, '');
      Exit;
    end;

    _BufferWriteIndex := (_BufferWriteIndex + 1) mod ((Integer(_BufferCount) + 1) * 2);
  except end;
end;

function TPlayer.GetOutputDeviceCount: Integer;
begin
  Result := waveOutGetNumDevs;
end;

function TPlayer.GetOutputDeviceName(index: Integer): String;
var
  outCaps: PWaveOutCaps;
  iErr: Integer;
begin
  outCaps := new(PWaveOutCaps);
  try
    iErr := waveOutGetDevCaps(index, outCaps, Sizeof(TWaveOutCapsA));

    if iErr <> 0 then begin
      Result := '';
      Exit;
    end;

    Result := outCaps^.szPname;

  finally
    Dispose(outCaps);
  end;
end;

procedure TPlayer.WMPaint(var Message: TWMPaint);
begin
  try
    if not _DrawTimmer.Enabled then begin
      DrawAudioLine(_MciAudioState = masStop);
    end;
  except end;

  inherited;
end;

procedure TAudioSettings.SetTitle(const Value: String);
begin
  _Title := Value;

  DrawAudioLine(_MciAudioState = masStop); 
end;

function TPlayer.GetStreamPosition: Int64;
begin
  if not Assigned(PlayStream) then begin
    Result := 0;
    Exit;
  end;
  
  Result := PlayStream.Position;
end;

procedure TPlayer.SetStreamPosition(const Value: Int64);
begin
  if not Assigned(PlayStream) then Exit;
  if PlayStream.Position >= PlayStream.Size then Exit;

  PlayStream.Position := value;
end;

procedure TAudioSettings.ShowFormatDialog;
var
  res: Longint;
begin
  if fc.pwfx = nil then
  begin

   fc.cbStruct := sizeof(fc);
   fc.pszTitle := '音频格式选择';
   fc.cbWfx := _WaveFmtSize;

   //getmem(fc.pwfx, MaxSizeFormat);
   {fc.pwfx.wFormatTag := pWaveFmt.wFormatTag;   //WAVE_FORMAT_GSM610; set default format to GSM6.10
   fc.pwfx.nChannels := pWaveFmt.nChannels;     //mono
   fc.pwfx.nSamplesPerSec := pWaveFmt.nSamplesPerSec;
   fc.pwfx.nAvgBytesPerSec:= pWaveFmt.nAvgBytesPerSec; // for buffer estimation
   fc.pwfx.nBlockAlign:= pWaveFmt.nBlockAlign;      // block size of data
   fc.pwfx.wbitspersample := pWaveFmt.wBitsPerSample;}

   fc.pwfx := pWaveFmt;

   //fc.szFormat := '48.000 kHz, 16 位, 立体声';
  end;

  fc.fdwStyle:=ACMFORMATCHOOSE_STYLEF_INITTOWFXSTRUCT;  //use the pwfx(waveformatex structure) as default
  res := acmFormatChoose(fc); //display the ACM dialog box

  if res=MMSYSERR_NOERROR then begin
    pWaveFmt := fc.pwfx;
    
    //FreeMemory;
    RefreshWaveFormat;
  end;
end;

function TAudioSettings.GetBPS: TBPS;
begin
  if pWaveFmt^.wBitsPerSample = C_BPS8 then
    Result := bps8
  else
    Result := bps16;
end;

function TAudioSettings.GetChannels: TChannels;
begin
  if pWaveFmt^.nChannels = C_MONO then
    Result := acMono
  else
    Result := acStereo;
end;

function TAudioSettings.GetSPS: Word;
begin
  Result := pWaveFmt^.nSamplesPerSec;
end;

function TAudioSettings.GetFormatTag: Integer;
begin
  Result := pWaveFmt^.wFormatTag;
end;

procedure TAudioSettings.SetFormatTag(const Value: Integer);
begin
  if pWaveFmt^.wFormatTag <> Value then begin
    pWaveFmt^.wFormatTag := Value;

    //FreeMemory;
    RefreshWaveFormat;
  end;
end;


{ TLameEnc }

constructor TLameEnc.Create;
begin
  _minBitRate := _BITRATE64;
  _maxBitRate := _BITRATE64;
  _avgBitRate := _BITRATE64;
  _private := False;
  _crc := True;
  _copyrighted := False;
  _original := False;
  _enableVBR := False;
  _vbrQuality := vbrQHigh;
  _disBRS := True;
  _lameSamples := 0;
  _useLame := True;
  _PlanRate := 0;
  _sampleRate := 44100;
  _channels := 2;

  _outBuf := nil;
  
  LoadLAME;
end;

destructor TLameEnc.Destroy;
begin
  if LameLoaded then UnloadLAME;
  
  inherited;
end;

function TLameEnc.InitStream(): Integer;
var
  nSamples: Cardinal;
  res: Integer;
begin
  Result := 0;
  
  if not _useLame then Exit;
  if not LameLoaded then Exit;
  
  fillChar(_lameConfig, sizeOf(_lameConfig), #0);
  _lameConfig.dwConfig := BE_CONFIG_LAME;

  with _lameConfig.r_lhv1 do begin
	  //Structure information
	  dwStructVersion := CURRENT_STRUCT_VERSION;
	  dwStructSize := sizeOf(_lameConfig);

    //Basic encoder setting
	  dwSampleRate := _sampleRate;  //SAMPLERATE OF INPUT FILE
	  dwReSampleRate := 0;  //DOWNSAMPLERATE, 0=ENCODER DECIDES  
	  nMode := {BE_MP3_MODE_STEREO;//{}_channels;//} //   OUTPUT   IN   STREO
   	dwBitrate := _minBitRate;        //set compress rate
   	dwMaxBitrate := _maxBitRate;
   	nPreset := LQP_HIGH_QUALITY;//{} LQP_NOPRESET;//{}LQP_R3MIX;//}LQP_PHONE;//} // Init   the   MP3   Stream
   	dwMpegVersion := {MPEG1;//{}MPEG2;//}
   	dwPsyModel := 0;  //   USE   DEFAULT   PSYCHOACOUSTIC   MODEL
   	dwEmphasis := 0;  //   NO   EMPHASIS   TURNED   ON

    //Bit Stream Settings
   	bPrivate := _private;
   	bCRC :=  _crc;
	  bCopyright := _copyrighted;
	  bOriginal :=  _original;

    //VBR Stuff    
   	bWriteVBRHeader := false;
   	bEnableVBR := _enableVBR;   //如果未开启vbr，则使用cbr方式固定码率
   	nVBRQuality := Integer(_vbrQuality);

   	if _enableVBR then begin
	    dwVbrAbr_bps := _avgBitRate;
	    if 0 < _avgBitRate then
	      nVbrMethod := VBR_METHOD_ABR
	    else
	      nVbrMethod := VBR_METHOD_NEW;

  	end	else
	    nVbrMethod := VBR_METHOD_NONE;

	  bNoRes := _disBRS;    // Disable Bit resorvoir (TRUE/FALSE)
	  bStrictIso := false;  // Use strict ISO encoding rules (TRUE/FALSE)
	  //btReserved := 0;  //FUTURE USE
  end;

  res := lameInitStream(PBE_CONFIG(@_lameConfig), nSamples, _minOutputBufSize, _hstream);
  if res <> BE_ERR_SUCCESSFUL then Exit;

  _lameSamples := nSamples shl 1; //这里必须乘以2，在调用lameEncodeChunk解码时，字节长度需要除以2
  mrealloc(_outBuf, _minOutputBufSize);

  if Assigned(_mp3Stream) then FreeAndNil(_mp3Stream);
  
  if Trim(_FileName) <> '' then
    _mp3Stream := TFileStream.Create(_FileName, fmCreate or fmShareDenyNone);
end;

function TLameEnc.CloseStream: Integer;
begin
  Result := 0;
  
  if not _useLame then Exit;
  if not LameLoaded then Exit;
  if _lameSamples <= 0 then Exit;

  result := lameDeinitStream(_hstream, _outBuf, _outBufUsed);
  //
  if (BE_ERR_SUCCESSFUL = result) then begin
    //
    // work around for stupid lame DLL bug
    lameWriteVBRHeader('');
    //
    //
    result := lameCloseStream(_hstream);
  end;

  if Assigned(_mp3Stream) then FreeAndNil(_mp3Stream);
  //
  _hstream := $FFFFFFFF;
end;

function TLameEnc.EncodeStream(data: pointer; const nBytes: Cardinal): Integer;
begin
  Result := 0;

  if data = nil then Exit;
  if nBytes = 0 then Exit;

  if LameLoaded and _useLame then begin
                                            
    result := lameEncodeChunk(_hstream, nBytes shr 1 {} {16 bits; regardless of number of channels}, data, _outBuf, _outBufUsed);
    if result <> BE_ERR_SUCCESSFUL then Exit;

    if Assigned(_mp3Stream) then _mp3Stream.Write(_outBuf^, _outBufUsed);
  end;

  Result := E_FAIL;  //未编码则返回false
end;

function TLameEnc.EncodeFile(const inputFile, outputFile: String): Boolean;
var
  temp:array[0..255] of byte;
  
  inputStream, outputStream: TFileStream;
  inputBuffer: PShortInt;
  readLen: Cardinal;
  i, res: Integer;
begin
  Result := False;

  if not FileExists(inputFile) then Exit;
  outputStream := TFileStream.Create(outputFile, fmCreate or fmShareDenyNone);
  try
    inputStream := TFileStream.Create(inputFile, fmOpenRead);
    inputStream.Position := 0;
    
    // Seek back to start of WAV file,
    // but skip the first 44 bytes, since that's the WAV header
    inputStream.Seek(36, soBeginning);
    inputStream.Read(temp, 4);

    //对于某些wav格式的文件，头文件长度不是固定的44字节
    i := 0;
    while ((temp[i]<>$64) or (temp[i+1]<>$61) or (temp[i+2]<>$74) or (temp[i+3]<>$61)) do begin
      inputStream.Read(temp[i+4],1);
      inc(i);
    end;       
    inputStream.Position := inputStream.Position + 4;
    
    try
      GetMem(inputBuffer, _lameSamples);

      try

        while true do begin
          readLen := inputStream.Read(inputBuffer^, _lameSamples);
          if readLen <= 0 then break;

          res := EncodeStream(inputBuffer, _lameSamples);
          if res <> BE_ERR_SUCCESSFUL then Break;

          _PlanRate := Round(inputStream.Position / inputStream.size * 100);

          //避免cpu被大量占用
          Sleep(1); 
          Application.ProcessMessages;
        end;
        
      finally
        FreeMem(inputBuffer);
      end;
    finally
      FreeAndNil(inputStream); 
    end;
  finally
    FreeAndNil(outputStream);
  end;
end;

procedure TLameEnc.LoadWavFormat(const wavFile: String);
var
  fStream: TFileStream;
  wavFormat: PWaveFormatEx;

  temp:array[0..255] of byte;
begin
  if not FileExists(wavFile) then Exit;

  GetMem(wavFormat, SIZEOF(TWaveFormatEx));
  try
    fStream := TFileStream.Create(wavFile, fmOpenRead);
    try
      fStream.Read(temp,22);
      fStream.Read(temp,2);

      //设置通道数，单声道为1，双声道为2
      wavFormat^.nChannels := temp[0];

      fStream.Read(temp,2);

      //设置采样率
      wavFormat^.nSamplesPerSec := temp[1]*256 + temp[0];

      fStream.Read(temp,8);
      fStream.Read(temp,2);

      //设置采样位数
      wavFormat^.wBitsPerSample := temp[0];

      _sampleRate := wavFormat^.nSamplesPerSec;
      _channels := wavFormat^.nChannels;
    finally
      FreeAndNil(fStream);
    end;
  finally
    FreeMem(wavFormat);
  end;    
end;


function TLameEnc.GetCurTime: Integer;
begin
  if not Assigned(_mp3Stream) then begin
    Result := 0;
    Exit;
  end;  

  Result := Round(_mp3Stream.Position / (_avgBitRate * 1024 / 8));
end;

function TLameEnc.GetStreamPostion: Int64;
begin
  if not Assigned(_mp3Stream) then begin
    Result := 0;
    Exit;
  end;

  Result := _mp3Stream.Position;
end;

function TLameEnc.GetStreamSize: Int64;
begin
  if not Assigned(_mp3Stream) then begin
    Result := 0;
    Exit;
  end;

  Result := _mp3Stream.Size;
end;

function TLameEnc.GetTimeLen: Integer;
begin
  if not Assigned(_mp3Stream) then begin
    Result := 0;
    Exit;
  end;

  Result := Round(_mp3Stream.Size / (_avgBitRate * 1024 / 8));
end;

procedure TLameEnc.SetCurTime(const Value: Integer);
begin
  if not Assigned(_mp3Stream) then begin
    Exit;
  end;

  _mp3Stream.Position := Round(Int64(value) * Int64(_avgBitRate) * Int64(1024) / Int64(8));
end;

procedure TLameEnc.SetStreamPostion(const Value: Int64);
begin
  if not Assigned(_mp3Stream) then begin
    Exit;
  end;

  _mp3Stream.Position := Value;
end;


procedure TRecorder.SetCompressMp3(const Value: Boolean);
begin
  _lame.UseLame := Value;
end;

function TRecorder.GetCompressMp3: Boolean;
begin
  Result := _lame.UseLame;
end;

procedure TLameEnc.SaveAsFile(const fileName: String);
var
  p: PChar;
  tmpStream: TFileStream;
  i: Integer;
begin
  if not Assigned(_mp3Stream) then Exit;
  
  tmpStream := TFileStream.Create(fileName, fmCreate or fmShareDenyNone);
  try

    GetMem(p, _mp3Stream.size);
    try
      _mp3Stream.Position := 0;
      _mp3Stream.ReadBuffer(p^, _mp3Stream.Size);
      _mp3Stream.Position := _mp3Stream.Size;

      tmpStream.Position := 0;
      tmpStream.WriteBuffer(p^, _mp3Stream.Size);
    finally
      FreeMem(p);
    end;

    i:=tmpStream.Size-8;    { size of file  }
    tmpStream.Position:=4;
    tmpStream.write(i,4);
    i:=i-$24;               { size of data   }
    tmpStream.Position:=40;
    tmpStream.write(i,4);
  finally
    FreeAndNil(tmpStream);
  end;  
end;


function TRecorder.GetMp3CompressRage: Integer;
begin
  Result := _lame._avgBitRate;
end;

procedure TRecorder.SetMp3CompressRate(const Value: Integer);
begin
  _lame._avgBitRate := Value;
  _lame._minBitRate := Value;
  _lame._maxBitRate := Value;
end;


procedure TAudioSettings.LoadDeviceSamples;
var
  curBlockAlign: Integer;
  curPerSec: Integer;
  curChunksPerSecond: Integer;
begin
  if not Assigned(pWaveFmt) then Exit;

  with pWaveFmt^ do begin
    nBlockAlign := (pWaveFmt^.wBitsPerSample * nChannels) shr 3 ;
    nAvgBytesPerSec := nSamplesPerSec * nBlockAlign;

    //参考voice communicatior 2.5 component   begin----------------------------
    if nBlockAlign = 0 then
      curBlockAlign := 512
    else
      curBlockAlign := nBlockAlign;  

    if (wFormatTag = WAVE_FORMAT_PCM) or (nAvgBytesPerSec = 0) then
      curPerSec := (max(1000, nSamplesPerSec) * max(Integer(1), nChannels) * max(Integer(8), wBitsPerSample)) shr 3
    else
      curPerSec := nAvgBytesPerSec;

    curChunksPerSecond := 25;

    _WaveBufSize := ((curPerSec div curChunksPerSecond + curBlockAlign - 1) div curBlockAlign) * curBlockAlign;

    //参考voice communicatior 2.5 component   End----------------------------
  end;
end;















procedure Register;
begin
  RegisterComponents('MCI', [TRecorder, TPlayer]);
end;





end.

