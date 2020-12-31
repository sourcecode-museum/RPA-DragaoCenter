program RPA;

uses
  Forms,
  uFormPrincipal in 'Forms\uFormPrincipal.pas' {formPrincipal},
  uFuncoes in 'Units\uFuncoes.pas',
  uFormSQLView in 'Forms\uFormSQLView.pas' {formSQLView},
  uFormSplash in 'Forms\uFormSplash.pas' {formSplash};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TformPrincipal, formPrincipal);
  Application.CreateForm(TformSplash, formSplash);
  Application.CreateForm(TformSQLView, formSQLView);
  Application.Title := 'Prestações em Atrazo';
  Application.Run;
end.
