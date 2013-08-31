unit MSRegExpr;

interface
uses VBScript_RegExp_55_TLB;
Type
  TMSRegExpr=class(TObject)
  private
    RegExprServer1:TRegExp;
    REInterface:OLEVariant;
    Text:WideString;
    ReplaceText:WideString;
    Procedure SetRegExprText(NewRegExpr:WideString);
    Procedure SetInputText(NewText:WideString);
    Procedure SetReplaceText(NewReplaceText:WideString);
    Procedure SetCaseSencitive(ACaseSens:Boolean);
  public
    Constructor Create;
    Destructor Destroy; override;
    Property RegExpr:WideString write SetRegExprText;
    Property InputText:WideString write SetInputText;
    Property ReplacementText:WideString write SetReplaceText;
    Property CaseSencitive:Boolean write SetCaseSencitive;
    Function ExecRegExpr:Boolean;
    Function ParcedText:WideString;
    Function ReplaceOnRequest(AInput,ARegExpr,AReplace:WideString;SaveNew:Boolean):WideString;
    Function Match(S,Pattern:WideString):Boolean;
  end;
implementation
Uses ActiveX,SysUtils;
Procedure TMSRegExpr.SetRegExprText(NewRegExpr:WideString);
Begin
  REInterface.Pattern:=NewRegExpr;
end;

Procedure TMSRegExpr.SetInputText(NewText:WideString);
Begin
  Text:=NewText;
end;

Procedure TMSRegExpr.SetReplaceText(NewReplaceText:WideString);
Begin
  ReplaceText:=NewReplaceText;
end;

Procedure TMSRegExpr.SetCaseSencitive(ACaseSens:Boolean);
Begin
  REInterface.IgnoreCase:= not ACaseSens;
end;

Function TMSRegExpr.ExecRegExpr;
Begin
  Result:=REInterface.Test(Text);
end;

Function TMSRegExpr.ParcedText:WideString;
Begin
  Result:=REInterface.Replace(TEXT,ReplaceText);
end;

Function TMSRegExpr.ReplaceOnRequest;
Begin
  RegExpr:=ARegExpr;
  ReplacementText:=AReplace;
  if AInput<>'' then
    InputText:=AInput;
  Result:=ParcedText;
  if SaveNew then
    InputText:=Result;
end;

Function TMSRegExpr.Match;
Begin
  RegExpr:=Pattern;
  if S<>'' then
    InputText :=S;
  result:=ExecRegExpr;
end;

Constructor TMSRegExpr.Create;
Begin
  CoInitialize(nil);
  RegExprServer1:=TRegExp.Create(Nil);
  REInterface:=RegExprServer1.DefaultInterface;
  REInterface.Global:=True;
end;

Destructor TMSRegExpr.Destroy;
Begin
  RegExprServer1.free;
  Inherited Destroy;
end;

end.
