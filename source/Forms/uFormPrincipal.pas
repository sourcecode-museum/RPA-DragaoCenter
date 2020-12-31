unit uFormPrincipal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  IBQuery, RxMemDS, IBSQL, IBDatabase, Db, IBCustomDataSet, Menus,
  ExtCtrls, Buttons, StdCtrls, CurrEdit, Mask, ToolEdit, Grids, DBGrids,
  DBCtrls, RXCtrls, FR_Desgn, FR_Shape, ImgList, FR_Class, DBTables,
  FR_DSet, FR_DBSet;

type
  TformPrincipal = class(TForm)
    pnlTop: TPanel;
    Image1: TImage;
    Label1: TRxLabel;
    pnlCenter: TPanel;
    Panel1: TPanel;
    Label4: TLabel;
    Label8: TLabel;
    Label12: TLabel;
    Bevel4: TBevel;
    DBGrid: TDBGrid;
    dt_PeriDe: TDateEdit;
    pnlTotal: TPanel;
    Panel3: TPanel;
    spdUnCheck: TSpeedButton;
    spdCheck: TSpeedButton;
    pnlBottom: TPanel;
    Shape1: TShape;
    lbl_Info: TLabel;
    MemoryPrinter: TRxMemoryData;
    ds_Memory: TDataSource;
    frReport1: TfrReport;
    ImgList_Grids: TImageList;
    frShapeObject1: TfrShapeObject;
    frDesigner1: TfrDesigner;
    dt_PeriAte: TDateEdit;
    Bevel2: TBevel;
    frDBDataSet1: TfrDBDataSet;
    Label2: TLabel;
    ed_ClieDe: TEdit;
    ed_ClieAte: TEdit;
    Label3: TLabel;
    ck_View: TCheckBox;
    spd_Printer: TSpeedButton;
    btn_Executar: TSpeedButton;
    Panel2: TPanel;
    DBNavigator1: TDBNavigator;
    Label5: TLabel;
    DBEdit2: TDBEdit;
    DBMemo1: TDBMemo;
    ed_Grupo: TEdit;
    Label6: TLabel;
    cmbOrdem: TComboBox;
    Bevel1: TBevel;
    Label7: TLabel;
    lblFiltrando: TLabel;
    Timer1: TTimer;
    lblCount: TRxLabel;
    QryPrincipal: TQuery;
    QryPrincipalDCODIC: TStringField;
    QryPrincipalDNOMEC: TStringField;
    QryPrincipalDENDERECO: TStringField;
    QryPrincipalDBAIRRO: TStringField;
    QryPrincipalDCIDADE: TStringField;
    QryPrincipalDUF: TStringField;
    QryPrincipalDCEP: TStringField;
    QryPrincipalDDATAPAGA: TDateField;
    QryPrincipalDVALORSIM: TFloatField;
    QryPrincipalDVALORAC1: TFloatField;
    QryPrincipalDVALORAC2: TFloatField;
    QryPrincipalDVALORDE1: TFloatField;
    QryPrincipalDVALORDE2: TFloatField;
    QryPrincipalDVALORTO1: TFloatField;
    QryPrincipalDVALORTO2: TFloatField;
    QryPrincipalDDOCUMENT: TStringField;
    QryPrincipalDGRUPO1: TStringField;
    QryPrincipalDCARTA: TStringField;
    DSPrincipal: TDataSource;
    TBLDizeres: TTable;
    ds_Dizeres: TDataSource;
    QryUpdateCarta: TQuery;
    ds_QryUpdateCarta: TDataSource;
    RxLabel1: TRxLabel;
    spd_Marcar: TSpeedButton;
    CheckBox1: TCheckBox;
    Label9: TLabel;
    dbRPA: TDatabase;
    procedure FormCreate(Sender: TObject);
    procedure spdCheckClick(Sender: TObject);
    procedure spdUnCheckClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure spd_PrinterClick(Sender: TObject);
    procedure frReport1GetValue(const ParName: String;
      var ParValue: Variant);
    procedure btn_ExecutarClick(Sender: TObject);
    procedure DBGridDblClick(Sender: TObject);
    procedure DBGridDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ed_ClieDeExit(Sender: TObject);
    procedure ed_ClieAteExit(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure cmbOrdemChange(Sender: TObject);
    procedure spd_MarcarClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);

  private
    { Private declarations }
    bInProcess : Boolean;
    mnComboItemIndex: integer;
    aUpdate : array of Variant;   //Array Cartas
    nUpdates : Integer;
    function OpenQryPrincipal(sOrderBy : string): boolean;
  public
    { Public declarations }
    function CampoOrderBy : string;
  end;

var
  formPrincipal: TformPrincipal;

implementation
  uses uFuncoes, uFormSQLView, uFormSplash;
{$R *.DFM}

function TformPrincipal.OpenQryPrincipal(sOrderBy : string):boolean;
var
  sSQL :string;
  nTime: TDateTime;
begin
  Result := False;
  Application.CreateForm(TformSplash, formSplash);
  formSplash.Show;
  formSplash.Update;

  try
    sSQL := 'SELECT DISTINCT(DCODIC), DNOMEC, DENDERECO, DBAIRRO, DCIDADE, DUF, DCEP,'
            + 'DGRUPO1, DDOCUMENT, DDATAPAGA, DVALORSIM, DVALORAC1, DVALORAC2,'
            + 'DVALORDE1, DVALORDE2, DVALORTO1, DVALORTO2, DCARTA'
            + ' FROM CTAREC.DBF'
            + ' INNER JOIN CLIENTE.DCB ON DCODIC = DCODIGO'
            + ' WHERE DDATAQUIT IS NULL';

    sOrderBy := Trim(sOrderBy);

    if sOrderBy <> '' then sSQL := sSQL + ' ORDER BY ' + sOrderBy ;

    with QryPrincipal do begin
      if Active then Close;

      SQL.Text := sSQL;

      Label9.Caption := '00:00:00';
      nTime := Now;
      Open;
      nTime := (Now - nTime);
      Label9.Caption := FormatDateTime('hh:mm:ss',nTime);

      lblCount.Caption := 'Quantidade de Registro: ' + IntToStr(RecordCount) + ' ';
    end;
    Result := True;
  except
    ShowMessage('Erro ao tentar abrir as TABELAS do Sistema SCA!'
                + #13 + 'Verifique o caminho do Banco de Dados no Arquivo Config.INI');
    Application.Terminate;
    Abort;
  end;

  formSplash.Hide;
  formSplash.Free;
end;

function TformPrincipal.CampoOrderBy:string;
begin
  if cmbOrdem.Text = 'Código' then
    Result := 'DCODIC'
  else if cmbOrdem.Text = 'Nome' then
    Result := 'DNOMEC'
  else if cmbOrdem.Text = 'Data Vencimento' then
    Result := 'DDATAPAGA'
  else
    Result := 'DCODIC,DDATAPAGA';
end;

procedure TformPrincipal.FormCreate(Sender: TObject);
begin
  nUpdates := 0;
  Sis_sPath := ExtractFilePath( Application.ExeName );
  Sis_sPathINI := Sis_sPath + 'Config.INI' ;

  if FileExists( Sis_sPathINI ) then begin
    Sis_sPathDB := Ler_INI( Sis_sPathINI, 'SISTEMA', 'PathDB');
    if Ler_INI(  Sis_sPathINI,'SISTEMA','Time') = 'S' then begin
      CheckBox1.Visible := True;
      Label9.Visible := True;
    end;
  end;

  if Sis_sPathDB = '' then begin
    Sis_sPathDB := Sis_sPath;
    Gravar_INI(Sis_sPathINI,'SISTEMA', 'PathDB',Sis_sPathDB);
  end;

  dt_PeriDe.Text := Ler_INI(Sis_sPathINI, 'PESQUISA', 'DataDe');
  dt_PeriAte.Text := Ler_INI(Sis_sPathINI, 'PESQUISA', 'DataAte');
  ed_ClieDe.Text := Ler_INI(Sis_sPathINI, 'PESQUISA', 'ClienteDe');
  ed_ClieAte.Text := Ler_INI(Sis_sPathINI, 'PESQUISA', 'ClienteAte');
  ed_Grupo.Text := Ler_INI(Sis_sPathINI, 'PESQUISA', 'Grupo');
  cmbOrdem.ItemIndex := StrToInt(Ler_INI(Sis_sPathINI, 'PESQUISA', 'Ordem'));
  mnComboItemIndex := cmbOrdem.ItemIndex; //Salvando valor antigo

  Sis_sRelEmpresa := Ler_INI(Sis_sPathINI, 'RELATORIO', 'Empresa');
  Sis_sRelSetor := Ler_INI(Sis_sPathINI, 'RELATORIO', 'Setor');
  Sis_sRelCidade := Ler_INI(Sis_sPathINI, 'RELATORIO', 'Cidade');
  Sis_sRelContato := Ler_INI(Sis_sPathINI, 'RELATORIO', 'Contato');

  //Setando ALias das Tabelas
  //QryUpdateCarta.DatabaseName := Sis_sPathDB;
  //QryPrincipal.DatabaseName := Sis_sPathDB;
  dbRPA.AliasName := Sis_sPathDB;
  dbRPA.Connected := True;
  TBLDizeres.DatabaseName := Sis_sPath;
  try
    TBLDizeres.Open;
  except
    MessageDlg('Tabela dos Dizeres não foi localizada!',mtError, [mbOK],0);
  end;
end;

procedure TformPrincipal.spdCheckClick(Sender: TObject);

begin
  try
    Label1.Visible := False;
    Image1.Visible := False;
    lblFiltrando.Visible := True;
    with QryPrincipal do begin
      if RecordCount <> 0 then begin
        First;

        bInProcess := True;
        while not Eof do begin
          if not bInProcess then begin
            Label1.Visible := True;
            Image1.Visible := True;
            lblFiltrando.Visible := False;
            Abort;
          end;

          DBGrid.SelectedRows.CurrentRowSelected := True;
          lblFiltrando.Caption := 'Selecionando Registros.' + #13
                        + IntToStr(DBGrid.SelectedRows.Count) + ' / ' + IntToStr(RecordCount);

          Next;
          Application.ProcessMessages;
        end;
      end;
    end;
  finally
    bInProcess := False;
  end;
  Label1.Visible := True;
  Image1.Visible := True;

  lblFiltrando.Visible := False;
  RxLabel1.Caption := 'Registros Selecionados: ' + IntToStr(DBGrid.SelectedRows.Count);
end;

procedure TformPrincipal.spdUnCheckClick(Sender: TObject);
var
  i,n,l : Integer;
begin
  i := 0; l:= 0;
  n := DBGrid.SelectedRows.Count;

  Label1.Visible := False;
  Image1.Visible := False;
  lblFiltrando.Visible := True;
  try
    with QryPrincipal do begin
      if (RecordCount <> 0) and (n > 0) then begin
        First;
        bInProcess := True;
        while not Eof and (i <= n) do begin
          if not bInProcess then begin
            Label1.Visible := True;
            Image1.Visible := True;
            lblFiltrando.Visible := False;

            Abort;
          end;

          if DBGrid.SelectedRows.CurrentRowSelected then begin
            Inc(i);
            DBGrid.SelectedRows.CurrentRowSelected:= False;
          end;
          inc(l);
          lblFiltrando.Caption := 'Desmarcando Seleção de Registros.' + #13
                        + IntToStr(RecordCount - l) + ' / ' + IntToStr(RecordCount);

          Next;
          Application.ProcessMessages;
        end; //While
      end; //if
    end; //With
  finally
    bInProcess := False;
  end;

  Label1.Visible := True;
  Image1.Visible := True;
  lblFiltrando.Visible := False;
  RxLabel1.Caption := 'Registros Selecionados: 0';
end;

procedure TformPrincipal.FormShow(Sender: TObject);
begin
  pnlCenter.Align := alClient;
  lblFiltrando.Visible := False;
  lblFiltrando.Align := alClient;

  dt_PeriDe.SetFocus;
end;

procedure TformPrincipal.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin  
  if Key = VK_ESCAPE then
    if bInProcess then
      bInProcess := False
    else
      Close
  else if (Shift = [ssCtrl]) and (Key = VK_F12) then
    formSQLView.show
  else if ( Shift = [ssCtrl] ) and ( Key = VK_F10 ) then
    frReport1.DesignReport
  else if spd_Printer.Enabled and (Key = VK_F10) then
    spd_PrinterClick( nil )
  else if btn_Executar.Enabled and (Key = VK_F9) then
    btn_ExecutarClick (nil)
  else if spd_Marcar.Enabled and (Key = VK_F11) then
    spd_MarcarClick (nil)

end;

procedure TformPrincipal.spd_PrinterClick(Sender: TObject);
var
  nSel,nCount,nDias : Integer;
begin
  lblFiltrando.Visible := True;
  lblFiltrando.Caption := 'Aguarde! Verificando dados para Impressão. ';
  Timer1.Enabled := True;

  nCount := 0;
  Screen.Cursor := crHourGlass;

  if MemoryPrinter.Active then MemoryPrinter.Close;

  SetLength(aUpdate,0); nUpdates := 0;//Limpando array antes de definir o tamanho

  nSel := DBGrid.SelectedRows.Count;
  if nSel >= 1 then QryPrincipal.First;

  SetLength(aUpdate,nSel); //definindo tamanho do array

  bInProcess := True;
  while not QryPrincipal.Eof and (nCount < nSel) do begin
    if not bInProcess then begin
      lblFiltrando.Visible := False;
      Timer1.Enabled := False;
      Abort;
    end;

    if DBGrid.SelectedRows.CurrentRowSelected then begin
      Inc(nCount);
      with MemoryPrinter do begin
        if not Active then Open;
        Append;
        FieldByName('DCODIC' ).AsString := QryPrincipal.FieldByName('DCODIC' ).AsString;
        FieldByName('DNOMEC' ).AsString  := QryPrincipal.FieldByName('DNOMEC' ).AsString;
        FieldByName('DENDERECO').AsString := QryPrincipal.FieldByName('DENDERECO').AsString;
        FieldByName('DBAIRRO').AsString := QryPrincipal.FieldByName('DBAIRRO').AsString;
        FieldByName('DCIDADE' ).AsString := QryPrincipal.FieldByName('DCIDADE').AsString;
        FieldByName('DUF' ).AsString := QryPrincipal.FieldByName('DUF').AsString;
        FieldByName('DDATAPAGA').AsDateTime  := QryPrincipal.FieldByName('DDATAPAGA').AsDateTime;
        FieldByName('DVALORSIM').AsFloat := QryPrincipal.FieldByName('DVALORSIM').AsFloat;
        FieldByName('DDOCUMENT' ).AsString := QryPrincipal.FieldByName('DDOCUMENT' ).AsString;

        nDias := DifDias(QryPrincipal.FieldByName('DDATAPAGA').AsDateTime,Date);
        FieldByName('vATRAZO').AsString := IntToStr(nDias) + ' dias';
      end; //with
    end;// SelectedRows
    QryPrincipal.Next;
  end; //while
  bInProcess := False;

  Timer1.Enabled := False;  //Para progresso de abetura

  if MemoryPrinter.IsEmpty then
    MsgErro( 'Nenhum registro foi marcado p/ impressão.' )
  else begin
    // Se continua faz a Impressão
    try
      lblFiltrando.Caption := 'Visualizando Impressão. ';

      frReport1.LoadFromFile( ExtractFilePath( Application.ExeName ) +
                            'CartaPrestacoesAtrazadas.frf' );
      if ck_View.Checked then
        frReport1.ShowReport
      else
        frReport1.PrintPreparedReportDlg;

      lblFiltrando.Visible := False;

      if MessageDlg('Desejá marcar cartas enviadas?',
                     mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        spd_MarcarClick(nil);

    finally
    end; //Try
  end;

  lblFiltrando.Visible := False;
  Timer1.Enabled := False;

  Screen.Cursor := crDefault;
end;

procedure TformPrincipal.frReport1GetValue(const ParName: String;
  var ParValue: Variant);

begin
  if ParName = 'vEMPRESA' then
    ParValue := Sis_sRelEmpresa
  else if ParName = 'vSETOR' then
    ParValue := Sis_sRelSetor
  else if ParName = 'vRODAPE' then
    ParValue := Sis_sRelContato
  else if ParName = 'vCIDADE' then
    ParValue := Sis_sRelCidade
  else if ParName = 'vDATAATUAL' then
    ParValue := FormatDateTime( 'dd" de "mmmm" de "yyyy, dddd ', Now )

  else if ParName = 'DCODIC' then
    ParValue := MemoryPrinter.FieldByName('DCODIC').AsString
  else if ParName = 'DNOMEC' then
    ParValue := MemoryPrinter.FieldByName('DNOMEC').AsString
  else if ParName = 'DENDERECO' then
    ParValue := MemoryPrinter.FieldByName('DENDERECO').AsString
  else if ParName = 'DBAIRRO' then
    ParValue := MemoryPrinter.FieldByName('DBAIRRO').AsString
  else if ParName = 'DCIDADE' then
    ParValue := MemoryPrinter.FieldByName('DCIDADE').AsString
  else if ParName = 'DUF' then
    ParValue := MemoryPrinter.FieldByName('DUF').AsString
  else if ParName = 'DCEP' then
    ParValue := MemoryPrinter.FieldByName('DCEP').AsString

end;

procedure TformPrincipal.btn_ExecutarClick(Sender: TObject);
var
  sDataDe, sDataAte: string;
  sOrdem, sFilter : string;
begin
  if not QryPrincipal.Active then
    if not Self.OpenQryPrincipal(Self.CampoOrderBy) then
      Exit;


  DBGrid.Enabled := True;
  DBGrid.SelectedRows.Clear;
  if MemoryPrinter.Active then MemoryPrinter.Close;

  if not DataValida(dt_PeriDe.Date) then begin
    MsgErro('Data inválida...');
    dt_PeriDe.SetFocus;
    Abort;
  end;

  if DataVazia(dt_PeriAte.EditText) then dt_PeriAte.EditText := dt_PeriDe.EditText;

  if dt_PeriDe.Date > dt_PeriAte.Date then begin
    MsgErro('Data inicial não pode ser maior que data Final...');
    Abort;
  end;

  lblFiltrando.Visible := True;
  lblFiltrando.Caption := 'Aguarde! Selecionando resgistros. ';
  Application.ProcessMessages;
  Timer1.Enabled := True;

  sDataDe := FormatDateTime('dd/mm/yyyy',dt_PeriDe.Date );
  sDataAte := FormatDateTime('dd/mm/yyyy',dt_PeriAte.Date );

  Screen.Cursor := crHourGlass;

  with QryPrincipal do begin
    Filtered := False;

    sFilter := 'DDATAPAGA >= ' + #39 + sDataDe + #39
             + ' AND DDATAPAGA <= ' + #39 + sDataAte + #39;

    if (Trim(ed_ClieDe.Text) <> '') and (Trim(ed_ClieDe.Text) <> '0000000') then  //Adicionado filtro para cliente
      sFilter := sFilter + ' AND DCODIC >= '+ #39 + ed_ClieDe.Text + #39;

    if (Trim(ed_ClieAte.Text) <> '') and (Trim(ed_ClieAte.Text) <> '0000000') then  //Adicionado filtro para cliente
      sFilter := sFilter + ' AND DCODIC <= ' + #39 + ed_ClieAte.Text + #39;

    if Trim(ed_Grupo.Text) <> '' then  //Adicionado filtro para cliente
      sFilter := sFilter + ' AND DGRUPO1 = ' + #39 + ed_Grupo.Text + #39;

    sOrdem := CampoOrderBy;

    {Win 2000-NT-XP}
    formSQLView.Memo1.Text := SQL.Text;
    //SQL.SaveToFile( 'C:\Documents and Settings\Administrador\Desktop\SQL.TXT' );
    {Win9x}
   // SQL.SaveToFile( 'C:\Windows\Desktop\SQL.txt' );

    try
      Application.ProcessMessages;
      Filter := sFilter;
      Filtered := True;

      lblFiltrando.Visible := False;
      Timer1.Enabled := False;

      if IsEmpty then begin
        lblCount.Caption := 'Quantidade de Registro: 0 ';
        MsgErro( 'Nenhum registro foi localizado p/ impressão.' );
        Close;
      end else begin
        lblCount.Caption := 'Quantidade de Registro: ' + IntToStr(RecordCount) + ' ';
        spdCheckClick(nil);
      end;

    except;
      ShowMessage('Erro ao tentar abrir esta Pesquisa!');
      Close;
    end; //Try
  end; //with

//  pnlTotal.Caption := 'Valor total em débitos: ' + FormatFloat('###,###,##0.00', fValorDebi);

  Screen.Cursor := crDefault;
end;

procedure TformPrincipal.DBGridDblClick(Sender: TObject);
begin
 if ( QryPrincipal.RecordCount > 0 ) then begin
    DBGrid.SelectedRows.CurrentRowSelected := not
                        DBGrid.SelectedRows.CurrentRowSelected;

    QryPrincipal.Next;
  end;

  RxLabel1.Caption := 'Registros Selecionados: ' + IntToStr(DBGrid.SelectedRows.Count);
end;

procedure TformPrincipal.DBGridDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  nMeio, nImag: Integer;

begin
  //nImag: 0 = Checked | 1 = UnChecked
  with DBGrid do begin
    if not SelectedRows.CurrentRowSelected then
      nImag := 1 //UnChecked
    else
      nImag := 0; //Checked

    if SelectedRows.Count = 0 then begin
      spd_Printer.Enabled := False;
      spd_Marcar.Enabled := False;
    end else begin
      spd_Printer.Enabled := DBGrid.Enabled;
      spd_Marcar.Enabled := DBGrid.Enabled;
    end;

    if DataCol = 0 then begin
      Canvas.Brush.Color := clWindow;
      Canvas.FillRect( Rect );

      nMeio := StrToInt( ( FormatFloat(
                           '##0', Columns[0].Width / 2 ) )) - 7;

      ImgList_Grids.Draw( Canvas, Rect.Left + nMeio, Rect.Top, nImag );
    end; //if DataCol
  end; //with
end;

procedure TformPrincipal.DBGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_SPACE then
    DBGridDblClick(nil);
end;

procedure TformPrincipal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Gravar_INI(Sis_sPathINI,'PESQUISA','DataDe', dt_PeriDe.Text);
  Gravar_INI(Sis_sPathINI,'PESQUISA','DataAte', dt_PeriAte.Text);
  Gravar_INI(Sis_sPathINI,'PESQUISA','ClienteDe', ed_ClieDe.Text);
  Gravar_INI(Sis_sPathINI,'PESQUISA','ClienteAte', ed_ClieAte.Text);
  Gravar_INI(Sis_sPathINI,'PESQUISA','Grupo', ed_Grupo.Text);
  Gravar_INI(Sis_sPathINI,'PESQUISA','Ordem', IntToStr(cmbOrdem.ItemIndex));

  try
    MemoryPrinter.Close;

    QryPrincipal.Close;
    TBLDizeres.Close;
    QryUpdateCarta.Close;
  finally
  end;
end;

procedure TformPrincipal.ed_ClieDeExit(Sender: TObject);
begin
  if ed_ClieDe.Text <> '' then
    ed_ClieDe.Text := FormatFloat('0000000', StrToInt(ed_ClieDe.Text));
end;

procedure TformPrincipal.ed_ClieAteExit(Sender: TObject);
begin
  if ed_ClieAte.Text <> '' then
    ed_ClieAte.Text := FormatFloat('0000000',StrToInt(ed_ClieAte.Text));
end;

procedure TformPrincipal.Timer1Timer(Sender: TObject);
begin
  lblFiltrando.Caption := lblFiltrando.Caption + '.';
  Application.ProcessMessages;
end;

procedure TformPrincipal.cmbOrdemChange(Sender: TObject);
begin
  if cmbOrdem.ItemIndex = mnComboItemIndex then
    Abort
  else begin
    if QryPrincipal.Active then
      if MessageDlg('O Banco de Dados deve ser Re-aberto.' + #13 +
                    'Esta operação pode demorar. Deseja continuar ?',
                     mtConfirmation,[mbYes, mbNo], 0) = mrYes then begin
        QryPrincipal.Close;
        mnComboItemIndex := cmbOrdem.ItemIndex;
      end else
        cmbOrdem.ItemIndex := mnComboItemIndex;
  end;
end;

procedure TformPrincipal.spd_MarcarClick(Sender: TObject);
  function StringUpdate(sDCarta,sDCodic, sDDocument: string;
                     pDDataPaga : TDateTime): string;
  var sWhere : string;
  begin
    Result := '';
    sWhere := 'DCODIGO = ' + #39 + sDCodic + #39
              + ' AND DDOCUMENT = ' + #39 + sDDOCUMENT + #39
              + ' AND DCARTA = ' + #39 + sDCarta + #39
              + ' AND DDATAPAGA = ' + #39
              + FormatDateTime('mm/dd/yyyy',pDDATAPAGA) + #39;

    if sDCarta = '' then
      sDCarta := '                '; //espaço de 16 caracteres

    if Length(sDCarta) <= (8 * 3) then begin //espaço para 3 datas
      sDCarta := sDCarta + FormatDateTime('yyyymmdd',Now);
      Result := 'UPDATE CTAREC.DBF SET DCARTA = '
         + #39 + sDCarta + #39
         + ' WHERE ' + sWhere;
    end;
  end;

  procedure SetControls(pEmProcesso : Boolean);
  begin
    lblFiltrando.Visible := pEmProcesso;
    Label1.Visible := not pEmProcesso;
    Image1.Visible := not pEmProcesso;

    btn_Executar.Enabled := not pEmProcesso;
    spd_Marcar.Enabled := not pEmProcesso;
    spd_Printer.Enabled := not pEmProcesso;
    DBGrid.Enabled := not pEmProcesso;
  end;

(* AQUI INICIA A PROCEDURE UPDATECARTAS *
 * ************************************ *)
var
  nSel,nCount,i : Integer;
  sStrUpdate: string;
const cCaption = 'Verificando Registros p/ marcar cartas! ';
begin
  Screen.Cursor := crHourGlass;

  i := 0; nCount := 0;

  nSel := DBGrid.SelectedRows.Count;
  if nSel >= 1 then begin
    QryPrincipal.First;
    SetControls(True);
    lblFiltrando.Caption := cCaption;
    Application.ProcessMessages;
  end;

  with QryPrincipal do begin
    bInProcess := True;
    while not Eof and (nCount < nSel) do begin
      if not bInProcess then begin
        SetControls(False);
        if nCount > 0 then begin
          QryUpdateCarta.Close;
          QryPrincipal.Close;
        end;

        Screen.Cursor := crDefault;
        Abort;
      end;

      if DBGrid.SelectedRows.CurrentRowSelected then begin
        Inc(nCount);
        sStrUpdate := StringUpdate(FieldByName('DCARTA').AsString,
                               FieldByName('DCODIC'    ).AsString,
                               FieldByName('DDOCUMENT' ).AsString,
                               FieldByName('DDATAPAGA' ).AsDateTime);
        if sStrUpdate <> '' then begin
          QryUpdateCarta.SQL.Text := sStrUpdate;
          QryUpdateCarta.ExecSQL;
        end;
      end;// SelectedRows

      Inc(i);
      lblFiltrando.Caption := cCaption + #13
            + 'Atualizados: ' + IntToStr(nCount) + ' / ' + IntToStr(nSel)
            + '  -  Qtd.Registros: ' + IntToStr(i) + ' / ' + IntToStr(RecordCount);
      Application.ProcessMessages;
      Next;
    end; //while
    bInProcess := False;
  end; //with

  if nCount > 0 then begin
    QryUpdateCarta.Close;
    QryPrincipal.Close;
  end;

  if nSel = 0 then
    MsgErro( 'Nenhum registro foi selecionado.' )
  else begin
    lblFiltrando.Visible := False;
    Label1.Visible := True;
    Image1.Visible := True;
    btn_Executar.Enabled := True;
  end;

  Screen.Cursor := crDefault;
end;

procedure TformPrincipal.CheckBox1Click(Sender: TObject);
begin
  QryUpdateCarta.RequestLive := CheckBox1.Checked;
  QryPrincipal.RequestLive := CheckBox1.Checked;
end;

end.

