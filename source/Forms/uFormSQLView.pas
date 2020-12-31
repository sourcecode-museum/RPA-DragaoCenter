unit uFormSQLView;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Buttons, ExtCtrls, Grids, DBGrids, StdCtrls, RXCtrls;

type
  TformSQLView = class(TForm)
    Memo1: TMemo;
    DBGrid1: TDBGrid;
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    SQLQuery: TQuery;
    ds_SQLQuery: TDataSource;
    Splitter1: TSplitter;
    lblCount: TRxLabel;
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Memo1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  formSQLView: TformSQLView;

implementation
  uses uFuncoes;
{$R *.DFM}

procedure TformSQLView.SpeedButton1Click(Sender: TObject);
var sSQL : string;
begin
  Screen.Cursor := crHourGlass;
  with SQLQuery do begin
    if Active then Close;
    Application.ProcessMessages;
    SSQL := Trim(MEMO1.Text);
    SQL.Text := SSQL;
    try
      if uppercase(Copy(sSQL,0,7)) = 'SELECT ' then begin
        Open;
        lblCount.Caption := 'Registros Encontrados: ' + IntToStr(RecordCount) + ' ';
      end else
        ExecSQL;
    finally
    end;
  end; //with
  Screen.Cursor := crDefault;
end;

procedure TformSQLView.FormCreate(Sender: TObject);
begin
  Self.Caption := 'SQLView -'
               + ' EMail: heliomarpm@hotmail.com / Celular: (033)8801-7394';

  DBGrid1.Align := alClient ;
end;

procedure TformSQLView.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  with SQLQuery do if Active then Close;
end;

procedure TformSQLView.Memo1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Shift = [ssCtrl]) and (Key = VK_RETURN) then
    SpeedButton1Click(nil);
end;

end.
