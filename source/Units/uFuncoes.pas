unit uFuncoes;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics,
  Controls, Forms, Dialogs, IniFiles;

  (* IniFiles *)
  function Ler_INI( pcArquivo, pcSecao, pcID: String ): String;
  procedure Gravar_INI( pcArquivo, pcSecao, pcID, pcValor: String );
  procedure MsgErro( cAviso: String );
  function DifDias(DataInicial, DataFinal:TDateTime): Integer;
  function DataVazia( pcData: String ): Boolean;
  function DataValida( const dData: TDateTime): Boolean;
var
  Sis_sPath   :string;
  Sis_sPathDB :string;
  Sis_sPathINI:string;

  Sis_sRelEmpresa :string;
  Sis_sRelSetor   :string;
  Sis_sRelCidade  :string;
  Sis_sRelContato :string;
implementation

{*************************************************************************
 ************************************************************************}
function Ler_INI( pcArquivo, pcSecao, pcID: String ): String;
(* Retorna o valor de um arquivo INI *)
var oArqINI: TIniFile;
begin
  Result := '';
  oArqINI := nil;

  try
    oArqINI := TIniFile.Create( pcArquivo );
    Result := oArqINI.ReadString( pcSecao, pcID, '' );
  finally
    oArqINI.Free;
  end;
end;

{*************************************************************************
 ************************************************************************}
procedure Gravar_INI( pcArquivo, pcSecao, pcID, pcValor: String );
(** Gravar em um arquivo INI *)
var oArqINI: TIniFile;
begin
  oArqINI := nil;
  try
    oArqINI := TIniFile.Create( pcArquivo );
    oArqINI.WriteString( pcSecao, pcID, pcValor );
  finally
    oArqINI.Free;
  end;                                      
end;

{*************************************************************
 ************************************************************}
procedure MsgErro( cAviso: String );
(* Função para dar avisos *)
begin
  MessageBeep(0);
  MessageDlg( cAviso, mtWarning, [mbOk], 0 );
end;

function DifDias(DataInicial, DataFinal:TDateTime): integer;
{Retorna a diferenca de dias entre duas datas}
var
  Data: TDateTime;
  dia, mes, ano: Word;
  sResult : string;
begin
  if DataFinal < DataInicial then begin
    Data := DataInicial - DataFinal;
    DecodeDate( Data, ano, mes, dia);
    sResult := '-' + FloatToStr(Data);
  end else begin
    Data := DataFinal - DataInicial;
    DecodeDate( Data, ano, mes, dia);
    sResult := FloatToStr(Data);
  end;

  Result := StrToInt(sResult);
end;

{*************************************************************************
 ** Esta função retorna se a data está vazia ou não, para isso ele deve **
 ** ser passada utilizando a propriedade EditText dos componentes des - **
 ** cendentes da class TDateTime                                        **
 ************************************************************************}
function DataVazia( pcData: String ): Boolean;
var
  pcDT: String[10];

begin
   pcDT := pcData;

   while Pos( '/', pcDT ) >0 do
     pcDT[ Pos( '/', pcDT ) ] := ' ';

   Result := ( Trim( pcDT ) = '' );
end;

{*************************************************************
 ** Faz a validação data e retorna verdadeiro ou falso      **
 *************************************************************}
function DataValida( const dData: TDateTime): Boolean;
var
  lRetorno: Boolean;
  nAno, nMes, nDia: Word;

begin
  try
    DateToStr( dData );
    lRetorno := True;
  except
    lRetorno := False;
  end;

  if lRetorno then begin
    DecodeDate( dData, nAno, nMes, nDia );
    if nAno <= 1899 then lRetorno := False;
  end;

  Result := lRetorno
end;
end.
