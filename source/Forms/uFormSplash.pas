unit uFormSplash;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, ImgList, RxGIF;

type
  TformSplash = class(TForm)
    IL: TImageList;
    Shape2: TShape;
    Label1: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    Shape3: TShape;
    Label4: TLabel;
    Label6: TLabel;
    Shape1: TShape;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  formSplash: TformSplash;

implementation
{$R *.DFM}

procedure TformSplash.FormCreate(Sender: TObject);
begin
 // IL.GetBitmap(1,Img_Logo.picture.bitmap);
end;

procedure TformSplash.FormShow(Sender: TObject);
begin
  SetWindowPos(Self.handle, HWND_TOPMOST,
               Self.Left, Self.Top,Self.Width, Self.Height, 0); // HWND_NOTOPMOST normal
end;

end.
