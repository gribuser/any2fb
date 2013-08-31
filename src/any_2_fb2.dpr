library any_2_fb2;

{%ToDo 'any_2_fb2.todo'}

uses
  ComServ,
  any_2_fb2_TLB in 'any_2_fb2_TLB.pas',
  Unit1 in 'Unit1.pas' {FBEImportPlugin: CoClass},
  any_2_fb_dialog in 'any_2_fb_dialog.pas' {UpenFileModalDialog},
  MSXML2_TLB in 'MSXML2_TLB.pas',
  DispatchForAny2fb in 'DispatchForAny2fb.pas',
  VBScript_RegExp_55_TLB in '..\..\..\Program Files\Borland\Delphi6\Imports\VBScript_RegExp_55_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}
{$R icon.res}
{R *.RES}

begin
end.
