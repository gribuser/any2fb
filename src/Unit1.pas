unit Unit1;


interface

uses
  Windows, ComObj, any_2_fb2_TLB, TX2FB, StdVcl;

type
  TFBEImportPlugin = class(TTypedComObject, IFBEImportPlugin)
  protected
    function import(hWnd: Integer; out filename: WideString;
      out document: IDispatch): HResult; stdcall;
  end;
  TOverridedComObjectFaactory=class(TTypedComObjectFactory)
    procedure UpdateRegistry(Register: Boolean); override;
  end;

implementation

uses ComServ,any_2_fb_dialog,Registry;

function TFBEImportPlugin.import(hWnd: Integer; out filename: WideString;
  out document: IDispatch): HResult;
Var
  FormResult:Integer;
begin
  UpenFileModalDialog:=TUpenFileModalDialog.GetDOMDocument(hWnd,FormResult,False);
  filename:=MyExtractFileName(UpenFileModalDialog.Edit1.Text);
  If Pos('.',FileName)<>0 then
    FileName:=Copy(FileName,1,Pos('.',FileName)-1);      
  Document:=UpenFileModalDialog.ResultDoc;
  UpenFileModalDialog.Free;
  If document<>Nil then
    Result:=S_OK
  else
    If FormResult<>idOk then
      Result:=S_FALSE
    else
      Result:=E_FAIL;
end;

Procedure TOverridedComObjectFaactory.UpdateRegistry;
Var
  FBEPluginKey:String;
  Reg:TRegistry;
  DLLPath:String;
Begin
  Inherited UpdateRegistry(Register);
  FBEPluginKey := 'SOFTWARE\Haali\FBE\Plugins\'+GUIDToString(CLASS_FBEImportPlugin);
  Reg:=TRegistry.Create;
  Try
    Reg.RootKey:=HKEY_LOCAL_MACHINE;
    If Register then
    Begin
      Reg.OpenKey(FBEPluginKey,True);
      Reg.WriteString('','"Any to FB2 v0.1" plugin by GribUser grib@gribuser.ru');
      Reg.WriteString('Menu','&ANY->FB2 by GribUser');
      Reg.WriteString('Type','Import');
      SetLength(DLLPath,2000);
      GetModuleFileName(HInstance,@DLLPath[1],1999);
      SetLength(DLLPath,Pos(#0,DLLPath)-1);
      Reg.WriteString('Icon',DLLPath+',0');
    end else Begin
      Reg.DeleteKey(FBEPluginKey);
      Reg.RootKey:=HKEY_CURRENT_USER;
      Reg.DeleteKey('software\Grib Soft\Any to FB2\1.0')
    end;
  finally
    Reg.Free;
  end;
end;

initialization
  TOverridedComObjectFaactory.Create(ComServer, TFBEImportPlugin, Class_FBEImportPlugin,
    ciMultiInstance, tmApartment);
end.
