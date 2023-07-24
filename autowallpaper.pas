unit autowallpaper;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,UrlMon,ShlObj, ComObj;

type
  TMinForm = class(TForm)
    btn1: TButton;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MinForm: TMinForm;

implementation

{$R *.dfm}

/// 用指定的方式改变墙纸

/// </summary>

/// <param name="strFile">墙纸图片文件</param>
/// <param name="style">样式</param>
procedure SetWallPaper(strFile: string; style: Integer);
var
　　dt : IActiveDesktop;
　　wpo : TWallPaperOpt;
　　ws : WideString;
begin
　　dt := CreateComObject(CLSID_ActiveDesktop) as IActiveDesktop;
　　ws := strFile;
　　case style of
　　　　0 : wpo.dwStyle := WPSTYLE_CENTER; //居中
　　　　1 : wpo.dwStyle := WPSTYLE_TILE; //平铺
　　　　2 : wpo.dwStyle := WPSTYLE_STRETCH; //拉伸
　　　　3 : wpo.dwStyle := WPSTYLE_MAX; //
　　else
　　　　wpo.dwStyle := WPSTYLE_CENTER;
　　end;
　　wpo.dwSize := SizeOf(wpo);
　　dt.SetWallpaperOptions(wpo, 0);
　　dt.SetWallpaper(PWideChar(ws), 0);
　　dt.ApplyChanges(AD_APPLY_ALL);
end;

  //文件下载
function DownloadFile(Source, Dest: string): Boolean;
begin
  try
    Result := UrlDownloadToFile(nil, PChar(source), PChar(Dest), 0, nil) = 0;
    except
      Result := False;
    end;
  end;

procedure TMinForm.FormCreate(Sender: TObject);
var filedir,downloadUrl1,downloadUrl2:string;
begin
    filedir :=ExtractFilePath(paramstr(0))+'desktop.jpg';
    downloadUrl1:='https://api.dujin.org/bing/1920.php';
    downloadUrl2:='https://www.yangshangzhen.com/bing/wallpaper';
    if not DownloadFile(downloadUrl1,filedir) then Exit;
    if not DownloadFile(downloadUrl2,filedir) then Exit;
    if FileExists(filedir) then
    SetWallPaper(filedir, 3)
    else
    ShowMessage('壁纸设置失败！请检查网络！');
    Application.Terminate;
end;

end.
