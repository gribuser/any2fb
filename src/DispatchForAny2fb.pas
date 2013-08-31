unit DispatchForAny2fb;


interface

uses
  ComObj, ActiveX, AxCtrls, any_2_fb2_TLB, StdVcl,tx2fb;

type
  TAny2FB2 = class(TAutoObject, IConnectionPointContainer, IAny2FB2)
  private
    { Private declarations }
    FConnectionPoints: TConnectionPoints;
    FConnectionPoint: TConnectionPoint;
    Params:TParceParams;
    Log:String;
  public
    procedure Initialize; override;
  protected
    { Protected declarations }
    property ConnectionPoints: TConnectionPoints read FConnectionPoints
      implements IConnectionPointContainer;
    function Convert(URL: OleVariant): IDispatch; safecall;
    function Get_noConvertCharset: WordBool; safecall;
    function Get_PreserveForm: WordBool; safecall;
    procedure Set_noConvertCharset(Value: WordBool); safecall;
    procedure Set_PreserveForm(Value: WordBool); safecall;
    function Get_noEpigraphs: WordBool; safecall;
    procedure Set_noEpigraphs(Value: WordBool); safecall;
    function Get_FixCount: Integer; safecall;
    function Get_noDescription: WordBool; safecall;
    function Get_noEmptyLines: WordBool; safecall;
    function Get_noFootNotes: WordBool; safecall;
    function Get_noItalic: WordBool; safecall;
    function Get_noQuotesConvertion: WordBool; safecall;
    procedure Set_FixCount(Value: Integer); safecall;
    procedure Set_noDescription(Value: WordBool); safecall;
    procedure Set_noEmptyLines(Value: WordBool); safecall;
    procedure Set_noFootNotes(Value: WordBool); safecall;
    procedure Set_noItalic(Value: WordBool); safecall;
    procedure Set_noQuotesConvertion(Value: WordBool); safecall;
    function ConvertInteractive(hWnd: Integer; needSave:WordBool): IDispatch; safecall;
    function Get_noPoems: WordBool; safecall;
    function Get_noRestoreBrokenParagraphs: WordBool; safecall;
    procedure Set_noPoems(Value: WordBool); safecall;
    procedure Set_noRestoreBrokenParagraphs(Value: WordBool); safecall;
    function Get_FollowLinksDeep: Integer; safecall;
    function Get_ignoreLineIndent: WordBool; safecall;
    function Get_leaveDinamicImages: WordBool; safecall;
    function Get_noExternalLinks: WordBool; safecall;
    function Get_noHeaders: WordBool; safecall;
    function Get_noImages: WordBool; safecall;
    function Get_noLongDashes: WordBool; safecall;
    function Get_noOffSiteImages: WordBool; safecall;
    procedure Set_FollowLinksDeep(Value: Integer); safecall;
    procedure Set_TextType(Value: Integer); safecall;
    procedure Set_ignoreLineIndent(Value: WordBool); safecall;
    procedure Set_leaveDinamicImages(Value: WordBool); safecall;
    procedure Set_noExternalLinks(Value: WordBool); safecall;
    procedure Set_noHeaders(Value: WordBool); safecall;
    procedure Set_noImages(Value: WordBool); safecall;
    procedure Set_noLongDashes(Value: WordBool); safecall;
    procedure Set_noOffSiteImages(Value: WordBool); safecall;
    function Get_FollowOffSiteLinks: WordBool; safecall;
    procedure Set_FollowOffSiteLinks(Value: WordBool); safecall;
    function Get_reHeadersDetect: OleVariant; safecall;
    function Get_reNeverFollowLinks: OleVariant; safecall;
    function Get_reOnDone: OleVariant; safecall;
    function Get_reOnLoad: OleVariant; safecall;
    function Get_reOnlyFollowLinks: OleVariant; safecall;
    procedure Set_reHeadersDetect(Value: OleVariant); safecall;
    procedure Set_reNeverFollowLinks(Value: OleVariant); safecall;
    procedure Set_reOnDone(Value: OleVariant); safecall;
    procedure Set_reOnLoad(Value: OleVariant); safecall;
    procedure Set_reOnlyFollowLinks(Value: OleVariant); safecall;
    function Get_LOG: OleVariant; safecall;

    Procedure InformAdded(S:String);
    Procedure WarningAdded(S:String);
  end;

implementation

uses ComServ,any_2_fb_dialog;

Procedure TAny2FB2.InformAdded(S:String);
Begin
  Log:=Log+S+#10;
end;
Procedure TAny2FB2.WarningAdded(S:String);
Begin
  Log:=Log+'Warning: '+S+#10;
end;

procedure TAny2FB2.Initialize;
begin
  inherited Initialize;
  FConnectionPoints := TConnectionPoints.Create(Self);
  if AutoFactory.EventTypeInfo <> nil then
    FConnectionPoint := FConnectionPoints.CreateConnectionPoint(
      AutoFactory.EventIID, ckSingle, EventConnect)
  else FConnectionPoint := nil;
  FillChar(Params,SizeOf(Params),0);
end;


function TAny2FB2.Convert(URL: OleVariant): IDispatch;
begin
  Log:='';
  Result:=Nil;
  Try
    Result:=ParseText(URL,@Params,InformAdded,WarningAdded,Nil);
  except
    Result:=Nil;
  end;
end;

function TAny2FB2.ConvertInteractive(hWnd: Integer;needSave:WordBool): IDispatch;
Var
  FormResult:Integer;
begin
  Log:='';
  UpenFileModalDialog:=TUpenFileModalDialog.GetDOMDocument(hWnd,FormResult,needSave);
  Result:=UpenFileModalDialog.ResultDoc;
  UpenFileModalDialog.Free;
{  If Result=Nil then
  Begin
    If FormResult<>mrOk then
      Result:=S_FALSE
    else
      Result:=E_FAIL;
  end;}
end;

function TAny2FB2.Get_noConvertCharset: WordBool;
begin
  Result:=Params.NotDetectEncoding;
end;

function TAny2FB2.Get_PreserveForm: WordBool;
begin
  Result:=Params.PreservForms;
end;

procedure TAny2FB2.Set_noConvertCharset(Value: WordBool);
begin
 Params.NotDetectEncoding:=Value;
end;

procedure TAny2FB2.Set_PreserveForm(Value: WordBool);
begin
 Params.PreservForms:=Value;
end;

function TAny2FB2.Get_noEpigraphs: WordBool;
begin
  Result:=Params.NotSearchEpig;
end;

procedure TAny2FB2.Set_noEpigraphs(Value: WordBool);
begin
 Params.NotSearchEpig:=Value;
end;

function TAny2FB2.Get_FixCount: Integer;
begin
  Result:=Params.FixTrialsOver100+100;
end;

function TAny2FB2.Get_noDescription: WordBool;
begin
  Result:=Params.NotSearchDescription;
end;

function TAny2FB2.Get_noEmptyLines: WordBool;
begin
  Result:=Params.Skip2Lines;
end;

function TAny2FB2.Get_noFootNotes: WordBool;
begin
  Result:=Params.NotDetectNotes;
end;

function TAny2FB2.Get_noItalic: WordBool;
begin
  Result:=Params.NotDetectItalic;
end;

function TAny2FB2.Get_noQuotesConvertion: WordBool;
begin
  Result:=Params.NotConvertQotes;
end;

procedure TAny2FB2.Set_FixCount(Value: Integer);
begin
  Params.FixTrialsOver100:=Value-100;
end;

procedure TAny2FB2.Set_noDescription(Value: WordBool);
begin
  Params.NotSearchDescription:=Value;
end;

procedure TAny2FB2.Set_noEmptyLines(Value: WordBool);
begin
  Params.Skip2Lines:=Value;
end;

procedure TAny2FB2.Set_noFootNotes(Value: WordBool);
begin
  Params.NotDetectNotes:=Value;
end;

procedure TAny2FB2.Set_noItalic(Value: WordBool);
begin
  Params.NotDetectItalic:=Value;
end;

procedure TAny2FB2.Set_noQuotesConvertion(Value: WordBool);
begin
  Params.NotConvertQotes:=Value;
end;

function TAny2FB2.Get_noPoems: WordBool;
begin
  Result:=Params.SkipPoems;
end;

function TAny2FB2.Get_noRestoreBrokenParagraphs: WordBool;
begin
  Result:=Params.NotRestoreParagr;
end;

procedure TAny2FB2.Set_noPoems(Value: WordBool);
begin
  Params.SkipPoems:=Value;
end;

procedure TAny2FB2.Set_noRestoreBrokenParagraphs(Value: WordBool);
begin
  Params.NotRestoreParagr:=Value;
end;

function TAny2FB2.Get_FollowLinksDeep: Integer;
begin
  Result:=Params.LinksDeep;
end;

function TAny2FB2.Get_ignoreLineIndent: WordBool;
begin
  Result:=Params.IgnoreSpaceAtStart;
end;

function TAny2FB2.Get_leaveDinamicImages: WordBool;
begin
  Result:=Params.KeepDynamicImages;
end;

function TAny2FB2.Get_noExternalLinks: WordBool;
begin
  Result:=Params.RemoveExternalLinks;
end;

function TAny2FB2.Get_noHeaders: WordBool;
begin
  Result:=Params.NoHeadDetect;
end;

function TAny2FB2.Get_noImages: WordBool;
begin
  Result:=Params.SkipAllImages;
end;

function TAny2FB2.Get_noLongDashes: WordBool;
begin
  Result:=Params.NotConvertDef;
end;

function TAny2FB2.Get_noOffSiteImages: WordBool;
begin
  Result:=Params.SkipOffSiteImages;
end;

procedure TAny2FB2.Set_FollowLinksDeep(Value: Integer);
begin
  Params.LinksDeep:=Value;
end;

procedure TAny2FB2.Set_TextType(Value: Integer);
begin
  Params.NoAutodetectFileType:=Value<>0;
  Params.HeadLowCence:=Value=2;
  If Params.NoAutodetectFileType then
    Params.ConvertLong:=80
  else
    Params.ConvertLong:=2;
end;

procedure TAny2FB2.Set_ignoreLineIndent(Value: WordBool);
begin
  Params.IgnoreSpaceAtStart:=Value;
end;

procedure TAny2FB2.Set_leaveDinamicImages(Value: WordBool);
begin
  Params.KeepDynamicImages:=Value;
end;

procedure TAny2FB2.Set_noExternalLinks(Value: WordBool);
begin
  Params.RemoveExternalLinks:=Value;
end;

procedure TAny2FB2.Set_noHeaders(Value: WordBool);
begin
  Params.NoHeadDetect:=Value;
end;

procedure TAny2FB2.Set_noImages(Value: WordBool);
begin
  Params.SkipAllImages:=Value;
end;

procedure TAny2FB2.Set_noLongDashes(Value: WordBool);
begin
  Params.NotConvertDef:=Value;
end;

procedure TAny2FB2.Set_noOffSiteImages(Value: WordBool);
begin
  Params.SkipOffSiteImages:=Value;
end;

function TAny2FB2.Get_FollowOffSiteLinks: WordBool;
begin
  Result:=Params.DownloadExternal;
end;

procedure TAny2FB2.Set_FollowOffSiteLinks(Value: WordBool);
begin
  Params.DownloadExternal:=Value;
end;

function TAny2FB2.Get_reHeadersDetect: OleVariant;
begin
  Result:=Params.HeaderMUSTRe;
end;

function TAny2FB2.Get_reNeverFollowLinks: OleVariant;
begin
  Result:=Params.LinksSkip;
end;

function TAny2FB2.Get_reOnDone: OleVariant;
begin
  Result:=Params.RegExpOnFinish;
end;

function TAny2FB2.Get_reOnLoad: OleVariant;
begin
  Result:=Params.REGExpOnStart;
end;

function TAny2FB2.Get_reOnlyFollowLinks: OleVariant;
begin
  Result:=Params.LinksFollow;
end;

procedure TAny2FB2.Set_reHeadersDetect(Value: OleVariant);
begin
  Params.HeaderMUSTRe:=Value;
end;

procedure TAny2FB2.Set_reNeverFollowLinks(Value: OleVariant);
begin
  Params.LinksSkip:=Value;
end;

procedure TAny2FB2.Set_reOnDone(Value: OleVariant);
begin
  Params.RegExpOnFinish:=Value;
end;

procedure TAny2FB2.Set_reOnLoad(Value: OleVariant);
begin
  Params.REGExpOnStart:=Value;
end;

procedure TAny2FB2.Set_reOnlyFollowLinks(Value: OleVariant);
begin
  Params.LinksFollow:=Value;
end;

function TAny2FB2.Get_LOG: OleVariant;
begin
  Result:=LOG;
end;

initialization
  TAutoObjectFactory.Create(ComServer, TAny2FB2, Class_Any2FB2,
    ciMultiInstance, tmApartment);
end.
