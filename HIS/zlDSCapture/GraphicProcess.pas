{*******************************************************************************
图像处理相关
创建人：TJH
创建日前：2009-11-10

描述：        
*******************************************************************************}

unit GraphicProcess;

interface

uses
  Windows, Classes, Sysutils, Messages, Graphics, Jpeg;


Type
  //图像处理类
  TGraphicProcess = class(Tobject)
  public
    //转换为灰度图
    class procedure ConvertBitmapToGrayscale(Bmp: TBitmap);
    //取得指定区域内的图像
    class procedure CutImg(const cutRect: TRect; sourceBmp, outBmp: TBitmap);
    //将BMP图像转换为JPG图像
    class function BmpConvertToJpg(sourceBmp: TBitmap; const compressRate: Integer): TJPEGImage;
  end;

implementation

class function TGraphicProcess.BmpConvertToJpg(
  sourceBmp: TBitmap; const compressRate: Integer): TJPEGImage;
var
  curRate: Integer;
begin
  try
    Result := TJPEGImage.Create;
    Result.Assign(sourceBmp);

    curRate := compressRate;
    if curRate > 100 then curRate := 100;
    if curRate < 0 then curRate := 0;

    Result.CompressionQuality := curRate;
    Result.Compress;
  except
    Result := nil;
  end;
end;

class procedure TGraphicProcess.ConvertBitmapToGrayscale(Bmp: TBitmap);
var
 x, y, Gray: Integer;
 Row: PRGBTriple;
 oldPixelFormat: TPixelFormat;
begin
  oldPixelFormat := Bmp.PixelFormat;
  Bmp.PixelFormat := pf24bit;

  for y := 0 to Bmp.Height - 1 do begin
    Row := Bmp.ScanLine[y];
    for x := 0 to Bmp.Width - 1 do begin
      Gray := (Row^.rgbtRed + Row^.rgbtGreen + Row^.rgbtBlue) div 3;
      Row^.rgbtRed := Gray;
      Row^.rgbtGreen := Gray;
      Row^.rgbtBlue := Gray;
      Inc(Row);
    end;
  end;

  Bmp.PixelFormat := oldPixelFormat;
end;

class procedure TGraphicProcess.CutImg(const cutRect: TRect; sourceBmp, outBmp: TBitmap);
begin
  outBmp.Width := cutRect.Right - cutRect.Left;
  outBmp.Height := cutRect.Bottom - cutRect.Top;

  outBmp.Canvas.CopyRect(Rect(0, 0, outBmp.Width, outBmp.Height), sourceBmp.Canvas, cutRect);
end;

end.
