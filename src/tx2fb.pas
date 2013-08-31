unit tx2fb;

interface
uses
  SysUtils,
  Windows,
  Classes,
  NMHttp,MSXML2_TLB,Graphics;

Type
  TInformProcedure=Procedure(S:String) of object;
  TProgressProcedure=Procedure(Percent:Integer) of object;
  TParceParams=record
    LoadFromString:Bool;
    NotDetectEncoding:Bool;
    NoAutodetectFileType:Bool;
    RemoveExternalLinks:Bool;
    SkipAllImages:Bool;
    SkipOffSiteImages:Bool;
    KeepDynamicImages:Bool;
    PreservForms:Bool;
    ConvertLong:Integer;
    HeadLowCence:Bool;
    NoHeadDetect:Bool;
    IgnoreSpaceAtStart:Bool;
    NotDetectItalic:Bool;
    NotDetectNotes:Bool;
    NotConvertQotes, NotConvertDef:Bool;
    Skip2Lines:Bool;
    NotSearchEpig:Bool;
    NotRestoreParagr:Bool;
    NotSearchDescription:Bool;
    FixTrialsOver100:Integer;
    SkipPoems:Bool;
    REGExpOnStart,RegExpOnFinish,HeaderMUSTRe,LinksFollow,LinksSkip:String;
    LinksDeep:Integer;
    DownloadExternal:Bool;
  end;
  PParceParam=^TParceParams;

  Function ParseText(AText:String;Params:PParceParam;AInformFunction,AWarningFunction:TInformProcedure;ACallBackFunction:TProgressProcedure):IXMLDOMDocument2; stdcall;
  Function MyExtractFileName(S:String):String;
implementation
uses MSRegExpr,NMUUE,Word97,pngimage,RXGif,FileCtrl;
Type
  TBigArray=Array [0..100] of Byte;
  PBigArray=^TBigArray;
Var
  Arr1,Arr2:TBigArray;
  PArr:PBigArray;
  FAutor,MAuthor,LAuthor,Title,ThisDocName:String;
  RegExp:TMSRegExpr;
  Description:String;
  InformFunction,WarningFunction:TInformProcedure;
  ProgressFunction:TProgressProcedure;
  PNGSteam:TMemoryStream;
  Deepness:Integer;

Const
  sUnformated=Nil;
  sTag=Pointer(1);
  sReady=Pointer(100);
  EntCount=100;
  Entities:Array[1..EntCount]of record Name:String;Sign:Char;
    end=((Name:'lt';Sign:#60),
      (Name:'gt';Sign:#62),
      (Name:'quot';Sign:#34),
      (Name:'nbsp';Sign:#160),
      (Name:'mdash';Sign:'—'),
      (Name:'iexcl';Sign:#161),
      (Name:'cent';Sign:#162),
      (Name:'pound';Sign:#163),
      (Name:'curren';Sign:#164),
      (Name:'yen';Sign:#165),
      (Name:'brvbar';Sign:#166),
      (Name:'sect';Sign:#167),
      (Name:'uml';Sign:#168),
      (Name:'copy';Sign:#169),
      (Name:'ordf';Sign:#170),
      (Name:'laquo';Sign:#171),
      (Name:'not';Sign:#172),
      (Name:'shy';Sign:#173),
      (Name:'reg';Sign:#174),
      (Name:'macr';Sign:#175),
      (Name:'deg';Sign:#176),
      (Name:'plusmn';Sign:#177),
      (Name:'sup2';Sign:#178),
      (Name:'sup3';Sign:#179),
      (Name:'acute';Sign:#180),
      (Name:'micro';Sign:#181),
      (Name:'para';Sign:#182),
      (Name:'middot';Sign:#183),
      (Name:'cedil';Sign:#184),
      (Name:'sup1';Sign:#185),
      (Name:'ordm';Sign:#186),
      (Name:'raquo';Sign:#187),
      (Name:'frac14';Sign:#188),
      (Name:'frac12';Sign:#189),
      (Name:'frac34';Sign:#190),
      (Name:'iquest';Sign:#191),
      (Name:'Agrave';Sign:#192),
      (Name:'Aacute';Sign:#193),
      (Name:'Acirc';Sign:#194),
      (Name:'Atilde';Sign:#195),
      (Name:'Auml';Sign:#196),
      (Name:'Aring';Sign:#197),
      (Name:'AElig';Sign:#198),
      (Name:'Ccedil';Sign:#199),
      (Name:'Egrave';Sign:#200),
      (Name:'Eacute';Sign:#201),
      (Name:'Eirc';Sign:#202),
      (Name:'Euml';Sign:#203),
      (Name:'Igrave';Sign:#204),
      (Name:'Iacute';Sign:#205),
      (Name:'Icirc';Sign:#206),
      (Name:'Iuml';Sign:#207),
      (Name:'ETH';Sign:#208),
      (Name:'Ntilde';Sign:#209),
      (Name:'Ograve';Sign:#210),
      (Name:'Oacute';Sign:#211),
      (Name:'Ocirc';Sign:#212),
      (Name:'Otilde';Sign:#213),
      (Name:'Ouml';Sign:#214),
      (Name:'times';Sign:#215),
      (Name:'Oslash';Sign:#216),
      (Name:'Ugrave';Sign:#217),
      (Name:'Uacute';Sign:#218),
      (Name:'Ucirc';Sign:#219),
      (Name:'Uuml';Sign:#220),
      (Name:'Yacute';Sign:#221),
      (Name:'THORN';Sign:#222),
      (Name:'szlig';Sign:#223),
      (Name:'agrave';Sign:#224),
      (Name:'aacute';Sign:#225),
      (Name:'acirc';Sign:#226),
      (Name:'atilde';Sign:#227),
      (Name:'auml';Sign:#228),
      (Name:'aring';Sign:#229),
      (Name:'aelig';Sign:#230),
      (Name:'ccedil';Sign:#231),
      (Name:'egrave';Sign:#232),
      (Name:'eacute';Sign:#233),
      (Name:'ecirc';Sign:#234),
      (Name:'euml';Sign:#235),
      (Name:'igrave';Sign:#236),
      (Name:'iacute';Sign:#237),
      (Name:'icirc';Sign:#238),
      (Name:'iuml';Sign:#239),
      (Name:'eth';Sign:#240),
      (Name:'ntilde';Sign:#241),
      (Name:'ograve';Sign:#242),
      (Name:'oacute';Sign:#243),
      (Name:'ocirc';Sign:#244),
      (Name:'otilde';Sign:#245),
      (Name:'ouml';Sign:#246),
      (Name:'divide';Sign:#247),
      (Name:'oslash';Sign:#248),
      (Name:'ugrave';Sign:#249),
      (Name:'uacute';Sign:#250),
      (Name:'ucirc';Sign:#251),
      (Name:'uuml';Sign:#252),
      (Name:'yacute';Sign:#253),
      (Name:'thorn';Sign:#254),
      (Name:'yuml';Sign:#255));
  EndZnak=['.','!','?','"',':','»','''','…'];
  headerSmartRE='';
//  HeadetMUSTRe='^\s*(Глава|Часть)';
  HeadDetectChars=['0'..'9','V','X','I','L','M','*','@'];
  LineChar=#150;
  IsInHTMLMode:Boolean=false;

Procedure Inform(InfoTExt:String);
Begin
  If Addr(InformFunction)<>Nil then
    InformFunction(InfoTExt);
end;
Procedure Warn(WarnText:String);
begin
  if Addr(WarningFunction)<>Nil then
    WarningFunction(WarnText);
end;

Procedure Progress(Percent:Integer);
Begin
  if Addr(ProgressFunction)<>Nil then
    ProgressFunction(Percent);
end;
Procedure Trim(Var S:String);
Const
  TrimSym=[' ',#9,#10,#13];
Begin
  if S='' then Exit;
  While (S<>'') and (S[1] in TrimSym) do
    Delete(S,1,1);
  While (S<>'') and (S[Length(S)] in TrimSym) do
    SetLength(S,Length(S)-1);
  While Copy(S,1,6)='&nbsp;' do
    Delete(S,1,6);
  While Copy(S,Length(S)-5,6)='&nbsp;' do
    SetLength(S,Length(S)-6);
end;
Function FTrim(S:String):String;
Begin
  trim(S);
  Result:=S;
end;
Function LineEmpty(S:String):Boolean;
Begin
  Result:=s='';
  If result then Exit;
  trim(S);
  Result:=(S='') or (s=#3) or (s='<p> &nbsp;</p>') or (s='<p>&nbsp;</p>') or (s='<empty-line/>');
end;
Function AllowedTag(Buf:String):Boolean;
Begin
  Result:=(Pos('<strong>',Buf)<>0) or (Pos('</strong>',Buf)<>0) or
    (Pos('<emphasis>',Buf)<>0) or (Pos('</emphasis>',Buf)<>0) or
    (Pos('<imagex',Buf)<>0) or (Pos('<a ',Buf)<>0) or (Pos('<id ',Buf)<>0)
    or (pos('<p>',Buf)<>0) or (pos('</x:a>',Buf)<>0) or
    (pos('<empty-line/>',Buf)<>0);
end;
Function IsDialChar(C:Char):Boolean;
Begin
  Result:= c in [#150,'-','-','—'];
end;

Function CleanTags(S:String):String;
Begin
  Result:=FTrim(RegExp.ReplaceOnRequest(S,'<[^>]*>','',False));
end;

Procedure ReplaceStr(var STR:String;SearchText,NewText:string);
Var
  FoundPos:Integer;
  NewStr:String;
Begin
  NewStr:='';
  FoundPos:=Pos(SearchText,STR);
  While FoundPos>0 do
  Begin
    NewStr:=NewStr+Copy(Str,1,FoundPos-1)+NewText;
    Str:=Copy(Str,FoundPos+Length(SearchText),MaxInt);
    FoundPos:=Pos(SearchText,STR);
  end;
  Str:=NewStr+Str;
end;
Function MyExtractFileName;
Begin
  Result:=RegExp.ReplaceOnRequest(S,'.*[\\\/]([^\\\/])(\?.*)?','$1',False);
  If Pos('?',Result)<>0 then
    Result:=Copy(Result,1,Pos('?',Result)-1);
  ReplaceStr(Result,'"','_');
  ReplaceStr(Result,'&','&amp;');
end;
Function DelEnt(S:String):String;
Var
  I,I1:Integer;
  Found:Boolean;
Begin
  I:=0;
  While I<Length(S) do
  Begin
    Found:=False;
    If (S[I]='&') and (Copy(S,I,4)<>'&lt;') and (Copy(S,I,5)<>'&amp;') and (Copy(S,I,4)<>'&gt;') and
    (Copy(S,I,2)<>'&#') then
    Begin
      For I1:=1 to EntCount do
        If copy(S,I+1,Length(Entities[I1].name)+1)=Entities[I1].name+';' then
        Begin
          Delete(S,I,Length(Entities[I1].name)+1);
          S[I]:=Entities[I1].sign;
          Found:=True;
          Break;
        end;
      If not Found then
        Insert('amp;',S,I+1);
    end;
    inc(I);
  end;
  Result:=S;
end;

FUnction WordToHTML(InDoc:String):String;
Var
  WordApp:_Application;
  Document:WordDocument;
  ConfirmConversions,ReadOnly,AddToRecentFiles,PasswordDocument,
  PasswordTemplate,Revert,WritePasswordDocument,WritePasswordTemplate,Format,
  FileFormat,LockComments,Password,WritePassword,ReadOnlyRecommended,
  EmbedTrueTypeFonts,SaveNativePictureFormat,SaveFormsData,
  SaveAsAOCELetter,SaveChanges,OriginalFormat,
  RouteDocument,FileName:OLEVariant;
  Ext:String;
  FN:String;
Const
  wdFormatHTML=8;
Begin
  WordApp := CoWordApplication.Create;
  Result:='';
  Try
    FileName:=InDoc;
    If Ext<>'.TXT' then
    Begin
      ConfirmConversions:=False;
      ReadOnly:=True;
      AddToRecentFiles:=False;
      PasswordDocument:='';
      PasswordTemplate:='';
      Revert:=False;
      WritePasswordDocument:='';
      WritePasswordTemplate:='';
      Format:=wdOpenFormatAuto;

      FileFormat:=wdFormatHTML;
      LockComments:=False;
      Password:='';
      WritePassword:='';
      ReadOnlyRecommended:=False;
      EmbedTrueTypeFonts:=False;
      SaveNativePictureFormat:=False;
      SaveFormsData:=False;
      SaveAsAOCELetter:=False;
      SaveChanges := WdDoNotSaveChanges;
//      OriginalFormat := UnAssigned;
 //     RouteDocument := UnAssigned;

      Document:=WordApp.Documents.Open(FileName,ConfirmConversions,ReadOnly,
                     AddToRecentFiles,PasswordDocument,PasswordTemplate,Revert,
                     WritePasswordDocument,WritePasswordTemplate,Format);
      SetLength(FN,501);
      GetTempPath(500,PChar(FN));
      SetLength(FN,Pos(#0,FN)-1);
      FileName:=FN+'$$Word_to_html_TMP_converter$$.tmp.html';
//      FileName:=RusToEng(Copy(FN,1,Length(FN)-Length(ExtractFileExt(FN))))+'.txt';
//      FileFormat:='wdFormatHTML';
      Document.SaveAs(FileName,FileFormat,LockComments,Password,
                     AddToRecentFiles,WritePassword,ReadOnlyRecommended,
                     EmbedTrueTypeFonts,SaveNativePictureFormat,SaveFormsData,
                     SaveAsAOCELetter);
      Document.Close(SaveChanges,OriginalFormat,RouteDocument);
      Result:=FileName;
      IsInHTMLMode:=true;
    end;
  Finally
    WordApp.Quit(SaveChanges,OriginalFormat,RouteDocument);
  end;
end;
Function FirstChar(S:String):Char;
// Возвращиет первый не-пробел.
Var
  I:Integer;
Begin
  For I:=1 to Length(S) do
    If S[I]<>' ' then
    Begin
      Result:=S[I];
      Exit;
    end;
  Result:=' ';
end;
Function PNGFromGif(InStream:TStream):TStream;
{Var
  CallSes:IPicture;
  HBitMap:THandle;
  Stream:IStream;
  BitMap:TBitMap;
  PNG:TPngObject;
  ThatW,ThatH:Integer;
  DrawRect:TRect;
  I,I1:Integer;
  MaskColor:TColor;
  ColorFound:Boolean;
  Buf:Array of byte;
  StreamSize:Integer;
  NewPos:Int64;
  CallRes:Integer;
Const}
//  PictureID:TGUID='{7BF80980-BF32-101A-8BBB-00AA00300CAB}';
{  HIMETRIC_INCH=2540;}
Var
  GIF:TGIFImage;
  PNG:TPngObject;
  BitMap:TBitMap;

Begin
{  Result:=Nil;
  CreateStreamOnHGlobal(0,True,Stream);
  StreamSize:=InStream.Seek(0,soFromEnd);
  SetLength(Buf,StreamSize);
  InStream.Seek(0,soFromBeginning);
  InStream.Read(Buf[1],StreamSize);

  Stream.Write(@Buf[1],StreamSize,@StreamSize);
  Stream.Seek(0,0,NewPos);
  CallRes:=OleLoadPicture(Stream,0,True,PictureID,CallSes);
  If not SUCCEEDED(CallRes) then Raise Exception.Create('AAA');
  InStream.Free;
  Stream._Release;

  CallSes.get_Handle(HBitmap);


  BitMap:=Graphics.TBitmap.Create;
    BitMap.PixelFormat:=pf24bit;
    CallSes.get_Width(ThatW);
    BitMap.Width:=Round(ThatW/HIMETRIC_INCH*GetDeviceCaps(BitMap.Canvas.Handle, LOGPIXELSX));
    CallSes.get_Height(ThatH);
    BitMap.Height:=Round(ThatH/HIMETRIC_INCH*GetDeviceCaps(BitMap.Canvas.Handle, LOGPIXELSY));
    DrawRect.Left:=0;
    DrawRect.Top:=0;
    DrawRect.Right:=BitMap.Width;
    DrawRect.Bottom:=BitMap.Height;
    CallSes.Render(BitMap.Canvas.Handle,0,0,DrawRect.Right-DrawRect.Left,
      DrawRect.Bottom-DrawRect.Top,0,ThatH,ThatW,-ThatH,DrawRect);
    MaskColor:=0;
    Repeat
      ColorFound:=True;
      For I:=0 to DrawRect.Right do
      Begin
        For I1:=0 to DrawRect.Bottom do
          If BitMap.Canvas.Pixels[I,I1]=MaskColor then
          Begin
            ColorFound:=False;
            Break;
          end;
        If not ColorFound then Break;
      end;
      Inc(MaskColor);
    until ColorFound;
    BitMap.Canvas.Brush.Color:=MaskColor;
    BitMap.Canvas.FillRect(DrawRect);
    CallSes.Render(BitMap.Canvas.Handle,0,0,DrawRect.Right-DrawRect.Left,
      DrawRect.Bottom-DrawRect.Top,0,ThatH,ThatW,-ThatH,DrawRect);
    CallSes._Release;}
    InStream.Seek(0,soFromBeginning);
    Gif:=TGIFImage.Create;
    Gif.LoadFromStream(InStream);
    InStream.Free;
//    Gif.SaveToFile('c:\temp\123.gif');
    Bitmap:=TBitmap.Create;
    Bitmap.Assign(Gif);
    PNG:=TPNGObject.Create;
    Try
      PNG.CompressionLevel:=9;
      PNG.Assign(Bitmap);
      png.RemoveTransparency;
//      PNG.TransparentColor:=MaskColor;
      PNGSteam:=TMemoryStream.Create;
      PNG.SaveToStream(PNGSteam);
//      PNG.SaveToFile('c:\temp\123.png');
    Finally
      PNG.Free;
    end;
   Gif.Free;
   Bitmap.Free;
end;
//======================================
Function ParseText;
Var
  FootNotes,UsedFiles,BInaryes:TStringList;
  XDoc:IXMLDOMDocument2;
  XMLHead:String;
  DonePart:String;
  IDPrefix:String;
  Done,Skipped:Integer;
  NedClean:String;
  WholeDoc:WideString;
  I:Integer;
  Ext:String;

  function GetTheText(Var AText:String;Params:PParceParam):TStringList;
  Var
    List:TStringList;
    HTTPSocket:TNMHTTP;
    InputString:String;
    TheFile:TFileStream;
    Buf:String;
    Ext:String;
    Function DecodeString(S:String):String;
    Var
      LooksLikeDOS:Integer;
      LooksLikeKoi:Integer;
      I:Integer;
      Procedure Coi8ToChar(Var S:String);
      Const
        Coi8:Array ['А'..'я'] of char='юабцдефгхийклмнопярстужвьызшэщчъЮАБЦДЕФГХИЙКЛМНОПЯРСТУЖВЬЫЗШЭЩЧЪ';
      Var
        I:Integer;
      Begin
        For I:=1 to Length(S) do
          If (S[I]>#191) then
          Begin
            S[I]:=Coi8[S[I]];
          end
          else
            Begin
              If S[i]=#$A3 then S[I]:='Ё' else
              Begin
                If S[I]=#$B3 then S[I]:='ё'
              end;
            end;
      end;
    Begin
      LooksLikeDOS:=0;
      If Pos(#160,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#161,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#162,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#167,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#173,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#160,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#173,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#149,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#151,S)>0 then
        Inc(LooksLikeDOS);
      If Pos(#175,S)>0 then
        Inc(LooksLikeDOS);

      LooksLikeKoi:=0;
      If Length(S)>2000 then
        I:=2000
      Else
        I:=Length(S);
        For I:=1000 to I do
          If (S[I]>='а') and (S[I]<='я') then
            Dec(LooksLikeKoi)
          else If (S[I]>='А') and (S[I]<='Я') then
            Inc(LooksLikeKoi);


      if Params.NotDetectEncoding then
      begin
        if (LooksLikeDOS>8) then
          Warn('Text seems to be MS-DOS RUS encoded, but text restoration was disabled!')
        else
          If (LooksLikeKoi>200) then
            Warn('Text seems to be KOI8-R encoded, but text restoration was disabled!');
        result:=S;
        Exit;
      end;
      If (LooksLikeDOS>8) then
      Begin
        OEMToChar(PChar(S),PChar(S));
        inform('MS-DOS rus encoding detected, converted to win-1251');
      end
      else
        If (LooksLikeKoi>200) then
        Begin
          Coi8ToChar(S);
          inform('KOI8-RUS encoding detected, converted to win-1251');
        end;
      result:=S;
    end;
    Procedure FindHeader(Str:String;Var FAutor,MAuthor,LAuthor,Title:String);
    Var
      I,TagPos:Integer;
      S,Autor,STitle:String;
      Buf:String;
      SpacePos:Integer;
      TitleClose:String;
      StartTagLen:Integer;
    Begin
      S:='';
      FAutor:='';
      MAuthor:='';
      LAuthor:='';
      Title:='';
      If Length(Str)<3000 then
        I:=Length(Str)-1
      else
        I:=3000;
      SetLength(S,3000);
      Move(Str[1],S[1],I);
      S:=LowerCase(S);
      StartTagLen:=1;
      TitleClose:=#$15;
      I:=Pos('<title>',S);
      if I<>0 then
      Begin
        TitleClose:='</title>';
        StartTagLen:=7;
      end else
        I:=Pos(#$14,S);

      If I<>0 then
      Begin
        STitle:=Copy(S,I+StartTagLen,Length(S)-I-StartTagLen);
        STitle:=Copy(STitle,1,Pos(TitleClose,STitle)-1);
        STitle:=Copy(Str,I+StartTagLen,Length(STitle));
        Title:=STitle;
        TagPos:=pos('<',Title);
        While TagPos<>0 do
        Begin
          Title[TagPos]:='&';
          Insert('lt;',Title,TagPos+1);
          TagPos:=pos('<',Title);
        end;
      end;

      I:=Pos('<meta name="author" content="',S);
      If I<>0 then
      Begin
        Autor:=Copy(Str,I+29,Length(S)-I-29);
        Autor:=Copy(Autor,1,Pos('"',Autor)-1);
      end else
      If (Title<>'') and (LAuthor='') and (pos('.',Title)<>0) then
      Begin
        Autor:=copy(Title,1,Pos('.',Title)-1);
        Title:=copy(Title,Pos('.',Title)+1,MaxInt);
      end;

      If Autor<>'' then
      Begin
        SpacePos:=Pos(' ',Autor);
        If SpacePos<>0 then
        Begin
          Buf:=Copy(Autor,1,SpacePos-1);
          Autor:=Copy(Autor,SpacePos+1,MaxInt);
          FAutor:=Buf;
          SpacePos:=Pos(' ',Autor);
          If SpacePos<>0 then
          Begin
            Buf:=Copy(Autor,1,SpacePos-1);
            Autor:=Copy(Autor,SpacePos+1,MaxInt);
            MAuthor:=Buf;
            LAuthor:=Autor;
          end else LAuthor:=Autor;
        end else LAuthor:=Autor;
      end;
      Trim(FAutor);Trim(MAuthor);Trim(LAuthor);Trim(Title);
    end;
    Procedure DetectFormat(Str:String);
    Var
      EndS,I:Integer;
      S:String;
      Lines:TStringList;
      AsHTM,AsFixed76,AsFixed80,AsLine,NoSpaces:Integer;
    Begin
      if Params.NoAutodetectFileType or (Length(str)<2000) then
        exit;

      If Length(Str)<10001 then
        EndS:=Length(Str)-1
      else
        EndS:=10000;
      S:=Copy(Str,2000,EndS-2000);
      Lines:=TStringList.Create;
      Try
        Lines.Text:=S;
        Lines.Delete(0);
        If Lines.Count<5 then
          Exit;
        Lines.Delete(Lines.Count-1);
        AsHTM:=0;
        AsFixed76:=0;
        AsFixed80:=0;
        AsLine:=0;
        NoSpaces:=0;
        For I:=0 to Lines.Count-1 do
        Begin
          S:=LowerCase(Lines[I]);
          If Pos('     ',S)=1 then
          Begin
            AsFixed76:=AsFixed76+2;
            AsFixed80:=AsFixed80+2;
            Dec(NoSpaces,10)
          end else
            Inc(NoSpaces);
          If Pos('              ',S)=1 then
            Dec(NoSpaces,100);
          If Pos('<pre',S)<>0 then
          Begin
            AsFixed76:=AsFixed76+10;
            AsFixed80:=AsFixed80+10;
          end;
          If (Length(s)<=80) and (Length(s)>76) then
            Inc(AsFixed80);

           If Length(s)=76 then
            Inc(AsFixed76);

          If Length(S)>80 then
            Inc(AsLine);
          If (Pos('<br>',S)<>0) or
            (Pos('<p',S)<>0) or
            (Pos('<div',S)<>0) then
            AsHTM:=AsHTM+2;
          if pos('<!doctype html',S)<>0 then
            AsHTM:=AsHTM+30;
        end;
      Finally
        Lines.Free;
      end;
      AsFixed80:=Round(AsFixed80/1.3);
      AsFixed76:=Round(AsFixed76/1.3);
      If (AsHTM>AsFixed80) and (AsHTM>AsFixed76) and (AsHTM>AsLine) then
      Begin
        Params.ConvertLong:=MaxInt;
        Params.IgnoreSpaceAtStart:=True;
        Inform('Text was recognized as html');
        IsInHTMLMode:=True;
      end else
      Begin
        Params.IgnoreSpaceAtStart:=False;
        If (AsLine>AsHTM) and (AsLine>AsFixed80) and (AsLine>AsFixed76) then
          Begin
            Params.ConvertLong:=1;
            Inform('Text was recognized as traditional TXT (line=paragraph)');
          end else
              if  AsFixed76>AsFixed80 then
              Begin
                Params.ConvertLong:=76;
                Inform('Text was recognized as fixed-width TXT (width=76)');
              end else
                Begin
                  Params.ConvertLong:=80;
                  Inform('Text was recognized as fixed-width TXT (width=80)');
                end;
        if NoSpaces>20 then
        Begin
          Params.HeadLowCence:=True;
          Inform('Headers detection set to LIGHT, some headers may be missed...');
        end;
      end;
    end;
    Function CleanUpText(Input:String):String;
    Var
      RegExpList:TStringList;
      I:Integer;
    Begin
      RegExp.InputText:=Input;
      ReplaceStr(Input,#$14,'<h1>');
      ReplaceStr(Input,#$15,'</h1>');

      if (Params.REGExpOnStart<>'') then
      Begin
        Inform('Running user regular expressions...');
        RegExpList:=TStringList.Create;
        RegExpList.Text:=Params.REGExpOnStart;
        If RegExpList.Count<=((RegExpList.Count-1) div 2)*2+1 then
          RegExpList.Add('');
        For I:=0 to (RegExpList.Count-1) div 2 do
        Begin
          Inform(#9'/'+RegExpList[I*2]+'/'+RegExpList[I*2+1]+'/');
          RegExp.ReplaceOnRequest('',RegExpList[I*2],RegExpList[I*2+1],True);
        end;
        RegExpList.Free;
      end;
      Inform('Preparing text for parsing...');
      Progress(8);
      if not params.SkipAllImages then
         RegExp.ReplaceOnRequest('','<\s*img[\s'#10#13'][^>]*?src=(["'']?)([^''">]+)\1[^>]*?>',#2712'$2'#2712,True);
      RegExp.ReplaceOnRequest('','<\s*a[\s'#10#13'][^>]*?name=(["'']?)([^''">]+)\1[^>]*?>',#2714'$2'#2714,True);
      Progress(10);
      RegExp.ReplaceOnRequest('','<\s*a[^>]*>[\s'#10#13']*</\s*a\s*>','',True);
      Progress(15);
      if not params.RemoveExternalLinks then
      Begin
        RegExp.ReplaceOnRequest('','<\s*a[\s'#10#13'][^>]*?href=(["'']?)'+ThisDocName+'#([^''">]+)\1[^>]*?>([\s\S'#10#13']*?)</\s*a\s*>',#2721'#$2'#2721'$3'#2721,True);
        RegExp.ReplaceOnRequest('','<\s*a[\s'#10#13'][^>]*?href=(["'']?)([^''">]+)\1[^>]*?>([\s\S'#10#13']*?)</\s*a\s*>',#2721'$2'#2721'$3'#2721,True);
      end else
        RegExp.ReplaceOnRequest('','<\s*a[\s'#10#13'][^>]*?href=(["'']?)'+ThisDocName+'#([^''">]+)\1[^>]*?>([\s\S'#10#13']*?)</\s*a\s*>',#2721'#$2'#2721'$3'#2721,True);
      Progress(20);
      RegExp.ReplaceOnRequest('','<\s*([/\w]+)[^>]*?>','<$1>',True);
      Progress(28);
      RegExp.ReplaceOnRequest('',#2712'([^'#2712']+?)'#2712,'<p><imagex xlink:href="$1"/></p>',True);
      Progress(35);
      RegExp.ReplaceOnRequest('',#2714'([^'#2714']+?)'#2714,'<id id="$1"/>',True);
      Progress(42);
      RegExp.ReplaceOnRequest('',#2721'([^'#2721']+?)'#2721'([^'#2721']*?)'#2721,'<a xlink:href="$1">$2</x:a>',True);
      Progress(48);
      RegExp.ReplaceOnRequest('','<a xlink:href="[^"]*[\[\{][^"]*">(.*?)</x:a>','$1',True);
      Progress(50);

      RegExp.ReplaceOnRequest('','<script>[\s\S'#10#13']*?</script>','',True);
      Progress(58);
      RegExp.ReplaceOnRequest('','</?(span|font|o:p)>','',True);
      RegExp.ReplaceOnRequest('','<![^>]*>','',True);
      Progress(65);
      RegExp.ReplaceOnRequest('','<(script|head)>[\s\S'#10#13']*?</\1>','',True);
      RegExp.ReplaceOnRequest('','</?(html|head|meta|body)>','',True);
      Progress(72);
      if not Params.PreservForms then RegExp.ReplaceOnRequest('','<form>[\s\S'#10#13']{0,2000}?</form>','',True);
      RegExp.ReplaceOnRequest('','<(/?)(i|sup|sub|em|code|tt)>','<$1emphasis>',True);
      Progress(80);
      RegExp.ReplaceOnRequest('','<(strong|emphasis)>([^<]*?)<\1>','<$1>$2',True);
      Progress(90);
      RegExp.ReplaceOnRequest('','<(/?)(b|strong)>','<$1strong>',True);
      result:=RegExp.ReplaceOnRequest('','</?(font|cpan)>','',True);
      Progress(100);
    end;
  Begin
    List:=TStringList.Create;
    if not Params.LoadFromString then
      if pos('http://',LowerCase(AText))=1 then
        Begin
          Buf:=AText;
          If Pos('#',Buf)<>0 then
          Begin
            SetLength(Buf,Pos('#',Buf)-1);
            AText:=PChar(Buf);
          end;

          Inform('Fetching URL via http...');
          HTTPSocket:=TNMHTTP.Create(Nil);
          Try
            HTTPSocket.OnPacketRecvd:=TNotifyEvent(ProgressFunction);
            HTTPSocket.Get(Buf);
            InputString:=HTTPSocket.Body;
            if Pos('text/html',LowerCase(HTTPSocket.Header))<>0 then
              IsInHTMLMode:=True;
          Finally
            HTTPSocket.Free;
          end;
          If Pos('?',AText)<>0 then
          Begin
            ThisDocName:=StrRScan(PChar(AText),'/');
            ReplaceStr(ThisDocName,'.','\.');
            ThisDocName:='(?:[^''"#>]*?'+Copy(ThisDocName,2,MaxInt)+')?';
          end else
            ThisDocName:=''
        end
      else
        Begin
          Inform('Loading data from file...');
          Ext:=UpperCase(ExtractFileExt(AText));
          NedClean:='';
          If ((Ext='.DOC') or (Ext='.DOT') or (Ext='.RTF') or (Ext='.WRI') or (Ext='.WK1') or
          (Ext='.WK3') or (Ext='.WK4') or (Ext='.MCW')) then
          Begin
            Inform('Using MSWord to convert document into html...');
            NedClean:=WordToHTML(AText);
            If NedClean='' then
              Raise Exception.Create('Unable to import document from MSWord file format!');
            ATExt:=PChar(NedClean);
          end else if not ((Ext='.TXT') or (Ext='.HTM') or (Ext='.HTML') or (Ext='.PRT')) then
              Warn('Unknown file extention!');
          if pos('HTM',Ext)<>0 then
            IsInHTMLMode:=True;
          TheFile:=TFileStream.Create(AText,fmOpenRead);
          Try
            If TheFile.Size=0 then
              Raise Exception.Create('Nothing to do:'#10'File is empty!!!')
            else If TheFile.Size<10 then
              Warn('File is WERY short. This may cause an error');
            SetLength(InputString,TheFile.Size);
            TheFile.Read(InputString[1],TheFile.Size);
          Finally
            TheFile.Free;
          end;
          Buf:=ExtractFileName(AText);
          ReplaceStr(Buf,'.','\.');
          ThisDocName:='(?:[^''"#>]*?'+Buf+')?';
        end
    else
      Begin
        InputString:=Copy(AText,1,MaxInt);
        ThisDocName:='[^''"#>]*?';
      end;

    InputString:=DecodeString(InputString);
    If UsedFiles.Count=0 then
      FindHeader(InputString,FAutor,MAuthor,LAuthor,Title);
    DetectFormat(InputString);
    InputString:=CleanUpText(InputString);
    List.Text:=InputString;
    GetTheText:=List;
    Inform('Text loaded OK, '+IntToStr(List.Count)+' lines total.');
  end;


//==============================================================================
//==============================================================================
Procedure RecogniseText(Var InputLines,NotesBody:TStringList);
// Самая главная функция, остальное - пользовательский интерфейс к ней.
  Function ValidHtml(Strings:TStringList):TStringList;
  //Разберем текст построчно, убрав лишние теги, табуляцию, комментарии и т.п.
  Var
    NewS:TStringList;
    CurLine,CurTag,Buf:String;
    I:Integer;
    Posit:Integer;
    OldI:Integer;
    FalsePosit:Integer;
    ComPos:Integer;
    MsgShown:Boolean;
    NotInHTMLStateNow:Integer;
    Function ClearTag(S:String):String;
    Begin
      While Pos(' <',S)<>0 do
        Delete(S,Pos(' <',S),1);
      While Pos('< ',S)<>0 do
        Delete(S,Pos('< ',S)+1,1);
      While Pos('> ',S)<>0 do
        Delete(S,Pos('> ',S)+1,1);
      While Pos(' >',S)<>0 do
        Delete(S,Pos(' >',S),1);
      While Pos('  ',S)<>0 do
        Delete(S,Pos('  ',S),1);
      Result:=LowerCase(S);
    end;
    Function RealyEmpty(S:String):Boolean;
    Begin
      trim(S);
      Result:=s='';
    end;
  Begin
    Inform('Marking html tags...');
    Progress(0);
    NewS:=TStringList.Create;
    Try
      NewS.Add('');
      NewS.Add('');
      I:=0;
      CurLine:='';
      oldI:=-1;
//      IsInHTMLMode:=False;
      NotInHTMLStateNow:=0;
      MsgShown:=False;
      While I< Strings.Count do
      Begin
        // Обратная связь о проделанной работе
        If Round((I/Strings.Count)*100)<>Round(((I-1)/Strings.Count)*100) then
          Progress(Round((I/Strings.Count)*100));
        // Если мы продвинулись на строку, не встретив ничего,
        // причешем новую строку
        If I>OldI then
        Begin
          If (Strings[I]<>'') and (CurLine<>'') then
            CurLine:=CurLine+' '+Strings[I]
          else
            CurLine:=CurLine+Strings[I];
          While Pos(#9,CurLine)<>0 do
          Begin
            Insert(' ',CurLine,Pos(#9,CurLine));
            CurLine[Pos(#9,CurLine)]:=' ';
          end;
          While Pos(#$A0,CurLine)<>0 do
            CurLine[Pos(#$A0,CurLine)]:=' ';
        end;
        OldI:=I;

        // Убираем комментарии
        ComPos:=Pos('<!--',CurLine);
        If ComPos<>0 then
        Begin
          If ComPos<>1 then
          Begin
            Strings.Insert(I,Copy(CurLine,1,ComPos-1));
            Strings[I+1]:=Copy(CurLine,ComPos,Length(CurLine));
            CurLine:=Copy(CurLine,1,ComPos-1);
            Continue;
          end;
          ComPos:=Pos('-->',CurLine);
          If ComPos<>0 then
          Begin
  // Не будем сохранять комметнарии
  //              NewS.AddObject(Copy(CurLine,1,ComPos+2),SComment);
            CurLine:=Copy(CurLine,ComPos+3,MaxInt);
            If CurLine='' then
              Inc(I);
            Continue;
          end;
          Inc(I);
          While (I<Strings.Count-1) and (Pos('-->',Strings[I])=0) do
          Begin
            CurLine:=CurLine+#10+Strings[I];
            Inc(I);
          end;
          ComPos:=Pos('-->',Strings[I]);
          If ComPos=Length(Strings[I])-2 then
          Begin
            CurLine:=CurLine+#10+Strings[I];
  //              NewS.AddObject(CurLine,SComment);
            CurLine:='';
            Inc(I);
            Continue;
          end else
            Begin
              CurLine:=CurLine+#10+Copy(Strings[I],1,ComPos+2);
  //                NewS.AddObject(CurLine,SComment);
              CurLine:=Copy(Strings[I],ComPos+3,Length(Strings[I]));
              Inc(I);
              Continue;
            end;
        end;

        // Так, теперь посмотрим, нет ли в строке тегов
        Posit:=Pos('<',CurLine);
        If Posit=0 then
        Begin
{          While pos(#13,CurLine)<>0 do
            CurLine[pos(#13,CurLine)]:=' ';
          While pos(#10,CurLine)<>0 do
            CurLine[pos(#10,CurLine)]:=' ';}
          if not ((NotInHTMLStateNow=0) and IsInHTMLMode) then
          Begin
            // Если в строке тегов нет, но мы разбираем HTML, то проигнорируем
            // перевод строки. Иначе - перенесем структкру в результат
            If RealyEmpty(CurLine) then
              NewS.AddObject('',sUnformated) else
              NewS.AddObject(CurLine,sUnformated);
            CurLine:='';
          end;
          Inc(I);
          Continue;
        end;
        If Posit<>1 then
        Begin
          // Если тег в строке встречается не в начале, скопируем начало к
          // прошлой строке в конец
          if (NotInHTMLStateNow=0) and IsInHTMLMode and (NewS.Objects[NewS.Count-1]=sUnformated) then
            NewS[NewS.Count-1]:=NewS[NewS.Count-1]+' '+Copy(CurLine,1,Posit-1)
          else
           Begin
             Buf:=Copy(CurLine,1,Posit-1);
             if (NotInHTMLStateNow=1) or not IsInHTMLMode or not LineEmpty(Buf) then
               NewS.AddObject(Buf,sUnformated);
           end;

          CurLine:=Copy(CurLine,Posit,Length(CurLine)-Posit+1);
          Continue;
        end;

        // Тэкс, имеем строку с тегом в начале...
        Posit:=Pos('>',CurLine);
        If Posit<>0 then
        Begin
          // Если конец тэга обнаружен, выкусим этот тег
          If ((Pos('</emphasis>',LowerCase(CurLine))=Posit-4) and (Posit<>4)) or
          ((Pos('</strong>',LowerCase(CurLine))=Posit-8) and (Posit<>8)) or
          ((Pos('</a>',LowerCase(CurLine))=Posit-3) and (Posit<>3)) then
          Begin
            FalsePosit:=Pos('<',Copy(CurLine,2,Length(CurLine)-1));
            If FalsePosit=0 then
              Begin
                NewS.AddObject(LowerCase(CurLine),sUnformated);
                Inc(I);
                CurLine:='';
                Continue;
              end Else
                Begin
                  NewS.AddObject(LowerCase(Copy(CurLine,1,FalsePosit)),sUnformated);
                  CurLine:=Copy(CurLine,FalsePosit+1,Length(CurLine)-FalsePosit+2);
                  Continue;
                end;
          end Else
          Begin
            CurTag:=ClearTag(Copy(CurLine,1,Posit));
            NewS.AddObject(CurTag,STag);
            if Pos('<pre',CurTag)=1 then
               Inc(NotInHTMLStateNow);
            if Pos('</pre',CurTag)=1 then
               Dec(NotInHTMLStateNow);
            CurLine:=Copy(CurLine,Posit+1,Length(CurLine)-Posit);
            If CurLine='' then Inc(I) Else
              While (CurLine<>'') and ((CurLine[1]=' ') or (CurLine[1]=#9)) do
                Delete(CurLine,1,1);
          end;
          Continue;
        end else
          If not MsgShown and (CurLine[1]='<') and (Length(CurLine)>1024) then
          Begin
            Warn('WERY long tag found. This probably is an error, large amount of text may be lost! Tag text:');
            Warn(Copy(CurLine,1,256));
            MsgShown:=True;
          end;

        Inc(I);
      end;
      If CurLine<>'' then
      Begin
        If CurLine[1]='<' then
          NewS.AddObject('&lt;'+Copy(CurLine,2,MaxInt),sUnformated)
        else
          NewS.AddObject(CurLine,sUnformated)

      end;
      Strings.Free;
      NewS.Add('');
      NewS.Add('');
      Result:=NewS;
    Except
      NewS.Free;
      Result:=Nil;
    end;
  end;

  Procedure RemoveFormsScripts(S:TStrings);
  Var
    I:Integer;
    DelLevel:Integer;
    DelCurrent:Boolean;
  Begin
    I:=1;
    DelLevel:=0;
    While I<S.Count do
    Begin
      DelCurrent:=False;
      If Pos('<script',S[I])<>0 then
        Inc(DelLevel);
      If Pos('<style',S[I])<>0 then
        Inc(DelLevel);
      If Pos('</style>',S[I])<>0 then
      Begin
        Dec(DelLevel);
        DelCurrent:=True;
      end;
      If Pos('</script>',S[I])<>0 then
      Begin
        Dec(DelLevel);
        DelCurrent:=True;
      end;

      If (DelLevel>0) or DelCurrent then
        S.Delete(I)
      else
        Inc(I);
    end;
  end;

  Procedure DetectParagraphs(S:TStringList);
  Var
    I:Integer;
    Str:String;
    Function IsParagraphDelimeter(S:String):Boolean;
    Begin
      S:=LowerCase(S);
      Result:=(S='') or (S='<center>') or (S='</div>') or (S='<div>') or
      (S='<dd>') or (S='</dd>') or (S='<dt>')or (S='</dt>') or
      (S='</center>') or(S='</p>') or (S='<p>') or
      (S='<br>') or (S='<br/>') or (Pos('</h',S)=1) or
      (Pos('</t',S)=1) or (Pos('<t',S)=1) or (Pos('<imagex',S)=1) or
      (Pos('</blockquote>',S)=1) or (Pos('<blockquote ',S)=1) or (Pos('<li>',S)=1);
    end;
    Function IsParagraphBeginer(S:String):Boolean;
    Begin
      Result:=(S='') or (S='<center>') or (S='<div>') or
      (S='<p>') or (S='<br>') or (S='<dd>') or (S='<dt>') or (S='<li>') or (S='<br/>') or
      (Pos('<t',S)=1) or
      (Pos('<blockquote ',S)=1);
    end;
    Procedure FindParagraphEnd(var I:Integer);
    Begin
      Repeat
        If (I>S.Count-3) then
          Break;
        If (S.Objects[I+1]<>sTag) then
          S[I+1]:=S[I]+' '+S[I+1]
        else
          if (AllowedTag(S[I+1])) and (pos('<p>',S[I+1])=0) then
            S[I+1]:=S[I]+S[I+1]
          else
            S[I+1]:=S[I];
        s.Objects[I]:=sReady;
        S[I]:=#3;
        Inc(I);
      Until IsParagraphDelimeter(S[I+1]);

      S[I]:=S[I]+'</p>';
      If (S[I]='<p></p>') and not Params.Skip2Lines then
        S[I]:='<empty-line/>';
      S.Objects[I]:=sReady;
    end;

  Begin
    Inform('Searching already HTML formating...');
    While I<S.Count-2 do
    Begin
      If s.Objects[I]<>sTag then
      Begin
        Inc(I);
        Continue;
      end;
      Str:=S[I];
{      If  (S[I]='<br>') then
        FindParagraphStart(I);}
      If (Pos('<h',S[I])<>0) then
      Begin
        S[I]:='</section><section><title><p>';
        FindParagraphEnd(I);
        S[I]:=S[I]+'</title>';
        While (I<S.Count-2) and ((LineEmpty(S[I+1])) or ((S.Objects[I+1]=sTag) and not (AllowedTag(S[I+1]) or (Pos('<h',S[I+1])<>0)))) do
        Begin
          s[I+1]:=S[I];
          S[I]:=#3;
          s.Objects[I]:=sReady;
          s.Objects[I+1]:=sReady;
          Inc(I);
        end;
        s[I-1]:=S[I];
        S[I]:=#3;
        Continue;
      end;
      If  IsParagraphBeginer(S[I]) then
      Begin
        if (Pos('<h',S[I+1])=0) then
        Begin
          S[I]:='<p>';
          FindParagraphEnd(I);
        end else S[I]:=#3;
        Inc(I);
        Continue;
      end;
      If not AllowedTag(S[I]) then
      Begin
        s.Objects[I]:=sReady;
//        s[I+1]:=s[I];
        S[I]:=#3;
      end else Inc(I);
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end;
  end;

  Procedure RemoveAllDSpaces(var S:TStringList);
  Var
    I,I1,PPos:Integer;
    Str,Buf:String;
    IDsCollected:TStringList;
    HRFound:String;
    PHref:PChar;
    newLines:TStringList;
  Const
    DeprecatedChars=#1#2#3#4#5#6#7#8#11#12#14#15#16#17#18#19#20#21#22#23#24#25#26#27#28#29#30#31;

    Function ClearS(Sub,S:String;I:Integer):String;
    Begin
      While Pos(Sub,S)>0 do
        Delete(S,Pos(Sub,S)+I,1);
      Result:=S;
    end;
    Function RegExpEscape(S:String):String;
    Var
      I:Integer;
    Begin
      I:=1;
      While I<=Length(S) do
      Begin
        If S[I] in ['(',')','[',']','\','*','.','^','?'] then
        Begin
          Insert('\',S,I);
          Inc(I);
        end;
        inc(I);
      end;
      Result:=S;
    end;
  Begin
    If S.Count=0 then
    Begin
      Warn('Empty text! This may be an error');
      Exit;
    end;
    If S[0]='<empty-line/>' then
      S[0]:='';
    IDsCollected:=TStringList.Create;
    Try
      Inform('Verifying internal links...');
      For I:=0 to S.Count-1 do
      Begin
        Buf:=S[I];
        If (Buf='') or (Buf=#3) then Continue;
        While Pos('<p> ',Buf)=1 do
          Delete(Buf,4,1);
        if (Buf='<p></p>') and not Params.Skip2Lines then
          Buf:='<empty-line/>';
        if (I<>0) and (Buf='<empty-line/>') and ((Pos('<section>',S[I-1])<>0) or (S[I-1]='<empty-line/>')
        or ((I<S.Count-2)and(pos('</section>',S[I+1])<>0))) then
        Begin
          Buf:='';
          Continue;
        End;

        if (pos('<p id="',Buf)<>0) and RegExp.Match(Buf,'<p id="([^"]*)"></p>') and (I<S.Count-2) then
        Begin
          Buf:=RegExp.ReplaceOnRequest(Buf,'<p id="([^"]*)"></p>','"$1"',True);
          I1:=I+1;
          While (I1<S.Count-1) do
          Begin
            PPos:=pos('<p>',S[I1]);
            if (PPos<>0) and (Pos('<p></p>',S[I1])=0) then
            Begin
              Str:=S[I1];
              Insert(' id='+Buf,Str,PPos+2);
              S[I1]:=Str;
              Break;
            end
            else Inc(I1);
          end;
          Buf:='<empty-line/>';
          RegExp.RegExpr:=Buf;
        end;

        If Pos('<id id=',Buf)<>0 then
        Begin
          RegExp.ReplaceOnRequest(Buf,'<p>(.*?)<id id="([^"]*?)"/>','<p id="$2">$1',True);
          While RegExp.Match('','<p id="([^"]*)[^\w-"]') do
            RegExp.ReplaceOnRequest('','<p id="([^"]*)[^\w-"]','<p id="$1_Q_',True);
          RegExp.ReplaceOnRequest('','<id id="[^"]*?"/>','',True);
          Buf:=RegExp.ReplaceOnRequest('','<p id="(\d)','<p id="fb_$1',True);
          HRFound:=Copy(Buf,8,Pos('">',Buf)-8);
          If RegExp.Match('','<p id="[^"]*?">\s*</p>') then
          Begin
            I1:=I+1;
            While (I1<S.Count-1) do
            Begin
              PPos:=pos('<p>',S[I1]);
              if (PPos<>0) and (Pos('<p></p>',S[I1])=0) then
              Begin
                Str:=S[I1];
                Insert(' id="'+HRFound+'"',Str,PPos+2);
                S[I1]:=Str;
                Break;
              end
              else Inc(I1);
            end;
            Buf:='<empty-line/>';
            RegExp.RegExpr:=Buf;
          end;
          If Pos('<p id="',Buf)=1 then
            if (IDsCollected.IndexOf(HRFound)=-1)then
              IDsCollected.Add(Copy(Buf,8,Pos('">',Buf)-8))
            else
              Buf:=RegExp.ReplaceOnRequest('','<p id="[^"]*">','<p>',True);
        end else
          If pos('<p><imagex xlink:href="',Buf)=1 then
            Buf:=Copy(Buf,4,Length(Buf)-7);
        If pos('<a xlink:href="#',Buf)<>0 then
        Begin
          RegExp.RegExpr:=Buf;
          While RegExp.Match('','<a xlink:href="#(\w*?)([^\w-"])') do
            RegExp.ReplaceOnRequest('','<a xlink:href="#([\w]*?)([^\w-"])','<a xlink:href="#$1_Q_',True);
          Buf:=RegExp.ReplaceOnRequest(Buf,'<a xlink:href="#(\d)','<a xlink:href="#fb_$1',True);
        end;
        S[I]:=Buf;
        RegExp.InputText:='';
        If Round((I/S.Count)*40)<>Round(((I-1)/S.Count)*40) then
          Progress(Round((I/S.Count)*40));
      end;
      For I:=1 to S.Count-1 do
      Begin
        If Round((I/S.Count)*60)<>Round(((I-1)/S.Count)*60) then
          Progress(Round((I/S.Count)*60)+40);
        Buf:=S[I];
        If (Buf='') or (Buf=#3) then Continue;
        PHref:=StrPos(@Buf[1],'<a xlink:href="#');
        While PHref<>Nil do
        Begin
          HRFound:=RegExpEscape(Copy(PHref,17,Pos('">',PHref)-17));
          If (IDsCollected.IndexOf(HRFound)=-1) and  (copy(HRFound,1,20)<>'FbAutId_')  then
          Begin
            if RegExp.Match(Buf,'<a xlink:href="#'+HRFound+'">(.*?)</x:a>') then
              Buf:=RegExp.ReplaceOnRequest('','<a xlink:href="#'+HRFound+'">(.*?)</x:a>','$1',True)
            else
              Buf:=RegExp.ReplaceOnRequest('','<a xlink:href="#'+HRFound+'">','',True);
            PHref:=StrPos(@Buf[1],'<a xlink:href="#');
          end else
            PHref:=StrPos(StrPos(PHref,'">'),'<a xlink:href="#');
        end;
      end;
    Finally
      IDsCollected.Free;
    end;
    For I:=1 to S.Count-1 do
    Begin
      For I1:=1 to Length(DeprecatedChars) do
        While Pos(DeprecatedChars[I1],S[I])<>0 do
        Begin
          Buf:=S[I];
          Buf[Pos(DeprecatedChars[I1],S[I])]:=' ';
          S[I]:=Buf;
        end;
    end;
    I:=0;
    NewLines:=TStringList.Create;
    Try
      Inform('Removing doblespaces...');
      While I< s.Count do
      Begin
        Str:=S[I];
        if pos('</section><section><title><p>Обращений с начала месяца:',S[I])=1 then
        Begin
          Inc(I);
          Continue;
        end;
        Str:=ClearS('  ',Str,0);
        Str:=ClearS('> <',Str,1);
        Str:=ClearS('> , <',Str,1);
        Str:=ClearS('> . <',Str,1);
        Str:=ClearS('> ! <',Str,1);
        Str:=ClearS('> ? <',Str,1);
        Str:=ClearS('< ',Str,1);
        Str:=ClearS(' >',Str,0);
        Str:=ClearS(' </p>',Str,0);
        Str:=ClearS('</x:a>',Str,2);
        Str:=ClearS('</:a>',Str,2);
        Str:=ClearS(' <p',Str,0);
        Str:=ClearS(' </strong>',Str,0);
        Str:=ClearS(' </emphasis>',Str,0);
        Str:=ClearS('<p> ',Str,3);
        Str:=ClearS('<p>'#160,Str,3);
        Str:=ClearS('<strong> ',Str,8);
        Str:=ClearS('<emphasis> ',Str,10);

        Str:=ClearS('--',Str,0);
        if (Str='<p></p>') then
        Begin
          if not Params.Skip2Lines then
            Str:='<empty-line/>'
          else
            Str:=#3;
        end;
        If Str<>S[I] then
        Begin
          S[I]:=Str;
          Continue;
        end;
        If not
          ((Str='') or
            (Str=#3) or
              (
                (NewLines.Count>0) and
                (Str='<empty-line/>') and
                (
                  (Pos('<section>',NewLines[NewLines.Count-1])<>0) or
                  (NewLines[NewLines.Count-1]='<empty-line/>') or
                  (
                    (I<S.Count-2) and
                    (pos('</section>',S[I+1])<>0)
                  )
                )
              )
            )
          then
            NewLines.Add(Str);
        Inc(I);
        If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
          Progress(Round((I/S.Count)*100));
      end;
    Finally
      S.Free;
      S:=NewLines;
    end;
  end;
  Procedure KillEntity(S:TStringList);
  Var
    I:Integer;
  Begin
    For I:=0 to S.Count-1 do
      If Pos('&',S[I])<>0 then
        S[I]:=DelEnt(S[I]);
  end;


  Procedure CreateParagraphs(S:TStringList);
  // А это - ядро функции, распознавание параграфов в тексте.
  Var
    I:Integer;

    Function MakeParagraphFromString(S:String):String;
    Begin
      Result:='<p>'+S+'</p>';
    end;
    Function NexTStringListhort:Boolean;
    Begin
      Result:=(S[I+1][Length(S[I+1])]in EndZnak) and (Length(S[I+1])<>Length(S[I]));
    end;

    Function ParInNextStr:boolean;
    Begin
      If I>=S.Count-1 then
        Result:= True
      else
      Begin
        result:=s.Objects[I+1]=sTag;
//        Str:=S[I+1];
//        Result:=(Pos('<p',Str)>0) or (pos('<section',Str)>0);
      end;
    end;

    Function IsDial(S:String):Boolean;
    Var
      I:Integer;
    Begin
      Result:=False;
      For I:=1 to Length(S) do
        If IsDialChar(S[I])then
        begin
          Result:=True;
          Exit;
        end else If S[I]<>' ' then Exit;
    end;
    Function NextLineStartsNew:Boolean;
    Begin
      Result:=(I>=S.Count-1) or (S.Objects[I+1]=sReady) or
        IsDial(S[I+1]) or ((pos('  ',S[I+1])=1) and not Params.IgnoreSpaceAtStart
        and ((Length(S[I])=0) or (S[I][Length(S[I])]in EndZnak) or Params.NotRestoreParagr))
        or LineEmpty(S[I+1]) or (Length(S[I+1])>Params.ConvertLong);
    end;
  Begin
    I:=0;
    Inform('Detecting TXT-styled paragraphs...');
    //Перебираем все строки с начала...
    While I<S.Count-2 do
    Begin
      If (S.Objects[I]=sReady) then
      Begin
        Inc(I);
        Continue;
      end;
      If NextLineStartsNew then
      Begin
        S[I]:=MakeParagraphFromString(S[I]);
        S.Objects[I]:=sReady;
        inc(I);
        If LineEmpty(S[I]) then
          S.Delete(I);
      end else
        Begin
          S[I]:=S[I]+' '+ S[I+1];
          S.Delete(I+1);
        end;
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end;
    For I:=0 to S.Count-1 do
    Begin
      If S.Objects[I]=sUnformated then
        S[I]:=MakeParagraphFromString(s[I]);
    end;
  end;

  Procedure RemoveNotClosed(S:TStringList);
  Var
    I:Integer;
    Buf:String;
    DeadPos:Integer;
    Procedure EncloseTags(var Buf:String;const TAG:String);
    Var
      FoundPos,ClosePos:PChar;
    Begin
      FoundPos:=StrPos(@Buf[1],PChar('<'+TAG));
      ClosePos:=StrPos(@Buf[1],PChar('</'+TAG));
      While (ClosePos<> Nil) and ((ClosePos<FoundPos) or (FoundPos=Nil))do
      Begin
        Move('<dd',ClosePos^,3);
        ClosePos:=StrPos(@Buf[1],PChar('</'+TAG));
      end;
      While FoundPos<>Nil do
      Begin
        ClosePos:=StrPos(FoundPos,PChar('</'+TAG));
        if ClosePos=Nil then
          Move('<dd',FoundPos^,3);
        FoundPos:=StrPos(ClosePos,PChar('<'+TAG));
      end;
      While ClosePos<>Nil do
      Begin
        ClosePos:=StrPos(StrPos(ClosePos,'>'),PChar('</'+TAG));
        if ClosePos<>Nil then
          Move('<dd',ClosePos^,3);
      end;
    end;
  Begin
    inform('Removing incorrect tags...');
    For I:=0 to S.Count-1 do
    Begin
      Buf:=S[I];
      EncloseTags(Buf,'emphasis');
      EncloseTags(Buf,'strong');
      EncloseTags(Buf,'a');
      DeadPos:=pos('<dd',Buf);
      While DeadPos>0 do
      Begin
        While (Buf[DeadPos]<>'>') and (DeadPos<Length(Buf)-1) do
          Delete(Buf,DeadPos,1);
        Delete(Buf,DeadPos,1);
        DeadPos:=pos('<dd',Buf);
      end;
      S[I]:=Buf;
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end;
  end;
  Procedure DetectHeaders(S:TStringList);
  Var
    I:Integer;
    LineEmptyI,LineEmpty1,LineEmpty2,LineEmpty3:Boolean;
    ThisIsIt:Boolean;
    Function IsRoman(S:String):Boolean;
    Begin
      Result:=False;
      if Pos('>',S)<>0 then
      Begin
        S:=Copy(S,Pos('>',S)+1,MaxInt);
        if Pos('<',S)<>0 then
          S:=Copy(S,1,Pos('<',S)-1);
      end;
      Trim(S);
      if Length(S)>8 then exit;
      Result:=RegExp.Match(S,'^m?m?m?(c[md]|d?c{0,3})(x[lc]|l?x{0,3})(i[xv]|v?i{0,3})$');
    end;
  Begin
    if params.NoHeadDetect and (Params.headerMustRE='') then Exit;
    I:=0;
    Inform('Searching for implicit headers...');
    While I<S.Count-5 do
    Begin
      If (LineEmpty(s[I]) and LineEmpty(S[I+1]) And (not LineEmpty(S[I+2]))
      and LineEmpty(S[I+3]) and (S.Objects[I+2]=Nil)) and not params.NoHeadDetect then
      Begin
        S[I+2]:='</section>'#10'<section><title><p>'+S[I+2]+'</p></title>'#10;

        S.Objects[I+2]:=sReady;
        S.Delete(I);
        S.Delete(I);
        While (I>0) and (I<S.Count-1) and (LineEmpty(S[I-1])) do
        Begin
          S.Delete(I-1);
          Dec(I);
        end;
        While (I<S.Count-2) and LineEmpty(S[I+1]) and LineEmpty(S[I+2]) do
          S.Delete(I+1);
      end Else
      Begin
        LineEmptyI:=LineEmpty(s[I]);
        if (I<S.Count-3) then
        Begin
          LineEmpty1:=LineEmpty(s[I+1]);
          LineEmpty2:=LineEmpty(s[I+2]);
          LineEmpty3:=LineEmpty(s[I+3]);
        end else
          Begin
            LineEmpty1:=False;
            LineEmpty2:=False;
            LineEmpty3:=False;
          end;
        ThisIsIt:=False;
        if (pos('<section>',S[I+1])=0) then
        Begin
          ThisIsIt:=not params.NoHeadDetect and
          (S.Objects[I+1]<>sReady) and
          (
            LineEmptyI and
            not LineEmpty1 and
            LineEmpty2 and
            (
                LineEmpty3 or
                (Pos('        ',S[I+1])<>0) or
                (S[I+1]=AnsiUpperCase(S[I+1])) or
                (
                  not IsDialChar(FirstChar(S[I+1])) and
                  not Params.HeadLowCence and
                  not (S[I+1][Length(S[I+1])] in EndZnak)
                ) or
                (S.Objects[I+1]=Nil) and
                (
                  (FirstChar(S[I+1]) in HeadDetectChars)
                ) or IsRoman(S[I+1])

            )
          );
          ThisIsIt:=ThisIsIt or
            not LineEmpty1 and
            (Params.headerMustRE<>'') and
            RegExp.Match(S[I+1],Params.headerMustRE);
        end;
        If ThisIsIt then
        Begin
          S[I+1]:='</section>'#10'<section><title><p>'+S[I+1]+'</p></title>'#10;
          S.Objects[I+1]:=sReady;
          S.Delete(I);
          While LineEmpty(S[I+1]){ and LineEmpty(S[I+2])} and (I<S.COunt-2) do
            S.Delete(I+1);
        end;
      end;
      Inc(I);
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end;
  end;

  Procedure DetectNesting(S:TStringList);
  Var
    I,I1,CurDeep,NewDeep:Integer;
  Begin
    CurDeep:=0;
    NewDeep:=0;
    if S.Count=0 then exit;
    Inform('Detecting sections nesting...');
    For I:=0 to S.Count-1 do
      If S[0]='<empty-line/>' then
        S.Delete(0)
      else
        Break;

{    if Copy(S[0],1,10)='</section>' then
      S[0]:=Copy(S[0],11,MaxInt)
    else
      S.Insert(0,'<section>');
    S.Add('</section>');}

    I1:=0;
    If (S.Count-1>100) and (not Params.NotSearchDescription) then
      For I:=0 to S.Count-2 do
      Begin
        If pos('<section>',S[I])<>0 then
        Begin
          if I>S.Count div 3 then Break;
          S[I]:=Copy(S[I],Pos('<section>',S[I])+9,MaxInt);
          I1:=I;
          Break;
        end;
      end{ else
        Exit};
    If Description='' then
      For I:=0 to I1-1 do
      Begin
        Description:=Description+S[0]+' ';
        S.Delete(0);
      end;
    S[0]:='<section>'+S[0];
    I:=0;
    While I<S.Count-2 do
    Begin
      If (Pos('<section>',S[I])<>0) and (Pos('epigraph>',S[I+1])<>0) then
      Begin
        S[I]:=S[I]+S[I+1];
        S.Delete(I+1);
      end else Inc(I);
    end;
    For I:=0 to S.Count-6 do
    Begin
      For I1:=1 to 5 do
        If (pos('<section>',S[I])<>0) and  (pos('</section>',S[I+I1])=1) then
        Begin
          If CurDeep>0 then
          Begin
            S[I]:='</section>'+S[I];
            Dec(CurDeep);
          end;
          S[I+I1]:=Copy(S[I+I1],11,MaxInt);
          inc(NewDeep);
        end else Break;
      CurDeep:=CurDeep+NewDeep;
      NewDeep:=0;
    end;
    For I:=0 to CurDeep do
      S[S.Count-1]:=S[S.Count-1]+'</section>';
  end;

  Procedure ItalicCreate(S:TStringList);
  Var
    I,I1:Integer;
    Buf:String;
    _Pos:Integer;
    FoundEnd:Boolean;
  Begin
    if Params.NotDetectItalic then Exit;
    inform('Searching for _italic_ text...');
    For I:=0 to S.Count-1 do
    Begin
      If (Pos('_',S[I])<>0) then
      Begin
        Buf:=S[I];
        RegExp.InputText:=S[I];
        While RegExp.Match('','(<[^>]*)_') do
          Buf:=RegExp.ReplaceOnRequest('','(<[^>]*)_','$1%5f',True);
        While Pos('_',Buf)<>0 do
        Begin
          _Pos:=Pos('_',Buf);
          Delete(Buf,_Pos,1);
          If Buf[_Pos-1]=' ' then
            Insert('<emphasis>',Buf,_Pos) else
            Insert('<emphasis>',Buf,_Pos-1);
          FoundEnd:=False;
          For I1:=_Pos+4 to Length(Buf) do
              If Buf[I1] in ['_','.',',','!','?',':','<'] then
                Begin
                  If Buf[I1]='_' then Delete(Buf,I1,1);
                  Insert('</emphasis>',Buf,I1);
                  FoundEnd:=True;
                  Break;
                end;
          If Not FoundEnd then
            Insert('</emphasis>',Buf,Pos('</p>',Buf)-1);
        end;
        Buf:=RegExp.ReplaceOnRequest(Buf,'\%5f','_',True);
        S[I]:=Buf;
      end;
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end
  end;

  Procedure DetectVerses(S:TStringList);
  Var
    I,I1,I2,FL:Integer;
    Buf,Clean:String;
  Begin
    If Params.SkipPoems then Exit;
    I:=0;
    Inform('Detecting verses...');
    While I< S.Count-5 do
    Begin
      Clean:=CleanTags(S[I]);
      If (Length(S[I])<80) and (Length(S[I+1])<80) and (Length(S[I+2])<80) and (Length(S[I+3])<80)
      and (pos('<section>',S[I])=0) and (Length(Clean)<60) and
      (S[I]<>'<empty-line/>') and (S[I+1]<>'<empty-line/>') and (S[I+2]<>'<empty-line/>')and
      (S[I+3]<>'<empty-line/>')then
      Begin
        FL:=Length(Clean);
        I1:=I;
        While I1<S.Count-1 do
          If (Abs(Length(Clean)-FL)>15) and (S[I1]<>'<empty-line/>') or
            (pos('<section>',S[I1])<>0) or IsDialChar(S[I1][Pos('>',S[I1])+1]) then
            Break
          else
            Inc(I1);
        If (I1-I>3) then
        Begin
          S.Insert(I,'<poem><stanza>');
          For I2:=I+1 to I1 do
          Begin
            If (S[I2]='<empty-line/>') then
              if (I2<>I1) then
                S[I2]:='</stanza><stanza>'
              else
                S[I2]:=''
            else
            Begin
              Buf:=S[I2];
              ReplaceStr(Buf,'<p','<v');
              ReplaceStr(Buf,'</p>','</v>');
              S[I2]:=Buf;
            end;
          end;
          S.Insert(I1+1,'</stanza></poem>');
          I:=I1+2;
        end;
      end;
      Inc(I);
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end;
  end;

  Procedure FormatQ(S:TStringList);
  Var
    I:Integer;
    Str:String;
  Begin
    If params.NotConvertQotes then Exit;

    For I:= 6 to S.Count-1 do
    Begin
      If (S.Objects[I] =sTag) or LineEMpty(S[I]) then
        Continue;
      Str:=S[I];
      If Length(Str)=0 then Continue;
      While Pos(' "',Str)<>0 do
        Str[Pos(' "',Str)+1]:='«';
      While Pos('-"',Str)<>0 do
        Str[Pos('-"',Str)+1]:='«';
      While Pos('("',Str)<>0 do
        Str[Pos('("',Str)+1]:='«';
      While Pos(';"',Str)<>0 do
        Str[Pos(';"',Str)+1]:='«';
      While Pos('" ',Str)<>0 do
        Str[Pos('" ',Str)]:='»';
      While Pos('"<',Str)<>0 do
        Str[Pos('"<',Str)]:='»';
      While Pos('"&',Str)<>0 do
        Str[Pos('"&',Str)]:='»';
      While Pos('".',Str)<>0 do
        Str[Pos('".',Str)]:='»';
      While Pos('",',Str)<>0 do
        Str[Pos('",',Str)]:='»';
      While Pos('";',Str)<>0 do
        Str[Pos('";',Str)]:='»';
      While Pos('":',Str)<>0 do
        Str[Pos('":',Str)]:='»';
      While Pos('"?',Str)<>0 do
        Str[Pos('"?',Str)]:='»';
      While Pos('"!',Str)<>0 do
        Str[Pos('"!',Str)]:='»';
      While Pos('")',Str)<>0 do
        Str[Pos('")',Str)]:='»';
      While Pos('"-',Str)<>0 do
        Str[Pos('"-',Str)]:='»';
      If Str[Length(Str)]='"' then
        Str[Length(Str)]:='»';
      If Str[1]='"' then
        Str[1]:='«';
      S[I]:=Str;
    end;
  end;


  Procedure DetectEpigraph(S:TStringList);
  Var
    I,I1:integer;
    EndFound:Integer;
    CoolLinesFound:Integer;
    FirstI:Integer;
    Ch:Char;

    Function FirstSIgnDial(S:String):Boolean;
    Begin
      If Length(S)=0 then
      Begin
        Result:=False;
        Exit;
      end;
      WHile (Length(S)>0) and (S[1]=' ') do
        Delete(S,1,1);
      Result:=(Length(S)>0) and ((IsDialChar(S[1])) or
      ((S[1] in ['0'..'9']) and (Length(S)>2) and (S[2] in [')','-','—',#150,' ','.',';'])));
    End;
  Label
    OnceAgain;
  Begin
    If params.NotSearchEpig then Exit;
    I:=2;
    While I<S.Count-20 do
    Begin
      if Pos('<section>',S[I])<>0 then
      Begin
        FirstI:=I;
        Repeat
          Inc(I)
        until (I<S.Count-2) or (S.Objects[i]=Nil);
        If (I-FirstI>16) or (S.Objects[I]=sTag)then Continue;
        If (I-FirstI=1) and not LineEmpty(S[I]) and (pos('         ',S[I])=0) then COntinue;
        OnceAgain:
        While (I<S.Count-2) and LineEmpty(S[i]) do
          Inc(I);
        I1:=I+60;
        If I1>S.Count-2 then
          I1:=S.Count-2;
        EndFound:=0;
        For I1:=I to I1 do
          If LineEmpty(S[I1]) or (Length(S[I1])>80) or (Length(S[I1])<5) or
          (S.Objects[I1]=sTag)then
          Begin
            EndFound:=I1;
            Break;
          end else
            if Pos('<section>',S[I1])<>0 then
            Begin
              EndFound:=I1-2;
              Break;
            end;
        If (EndFound>0) and (EndFound<>I)  then
        Begin
          //Если пустая строка найдена сравнительно недалеко, проверяем,
          //похож ли текст на эпиграф
          CoolLinesFound:=0;
          For I1:=I to EndFound do
          Begin
            If (Pos('          ',S[I1])<>0) or (Length(Ftrim(S[I1]))<60) and
              (not LineEmpty(S[I1])) then
              Inc(CoolLinesFound);
            If (Length(FTrim(S[I1]))>60) then
              Dec(CoolLinesFound,5);
            If FirstSIgnDial(S[I1]) then
              Dec(CoolLinesFound);
          end;
          If (EndFound=I) or (CoolLinesFound/(EndFound-I)>0.8) then
          //И если похоже - все строки вправо
          Begin
            S[I]:='<epigraph><p>'+S[I];
            For I1:=I to EndFound-3 do
            Begin
              if S[I+1]<>#3 then
                S[I]:=S[I]+' '+S[I+1];
              S.Delete(I+1);
            end;
            Ch:=FirstChar(S[I+1]);
            If (not LineEmpty(S[I+1])) and (AnsiUpperCase(Ch)=Ch) and not (S.Objects[I+1]=sTag) then
              S[I]:=S[I]+'</p><text-author>'+S[I+1]+'</text-author></epigraph>'
            else
              S[I]:=S[I]+' '+S[I+1]+'</p></epigraph>';
            S.Delete(I+1);

            S.Objects[I]:=sReady;
            While LineEmpty(S[I-1]) do
            Begin
              S.Delete(I-1);
              Dec(I);
            end;
            while (I<S.Count-1) and not (LineEmpty(S[i])) do
              Inc(I);
            If (Pos('      ',S[I+1])<>0) or (Length(S[I+1])<60) then
              Goto OnceAgain;
          end;
        end;
      end;
      Inc(I);
    end;
  end;
  Procedure FormatT(S:TStringList);
  Var
    I:Integer;
    Str:String;
    DialPos:Integer;
  Begin
    If Params.NotConvertDef then Exit;
    Inform('Fixing dialogs...');
    For I:= 0 to S.Count-1 do
    Begin
      Str:=S[I];
      While Pos('- ',Str)<>0 do
        Str[Pos('- ',Str)]:=LineChar;
      While Pos(' -',Str)<>0 do
        Str[Pos(' -',Str)+1]:=LineChar;
      DialPos:=Pos('>'+LineChar+' ',Str);
      If (DialPos<>0) and (DialPos=Pos('>',Str)) then
        Str[DialPos+2]:=#160;
      S[I]:=Str;
      If Round((I/S.Count)*100)<>Round(((I-1)/S.Count)*100) then
        Progress(Round((I/S.Count)*100));
    end;
  end;

  Procedure FormatP(S:TStringList);
  Var
    I,Posit:Integer;
    Str:String;
  Begin
    For I:= 0 to S.Count-1 do
    Begin
      Str:=S[I];
      While Pos('...',Str)<>0 do
      Begin
        Posit:=Pos('...',Str);
        Delete(Str,Pos('...',Str),2);
        Str[Posit]:=#133;
        S[I]:=Str;
      end;
    end;
  end;

  Procedure CreateFootnotes(S:TStringList);
  Var
    I:Integer;
    ID:Integer;
    Procedure CheckBlock(C1,C2:Char);
    Var
      Buf:String;
      Posit,ClosePosit:Integer;
    Begin
      Posit:=Pos(C1,S[I]);
      ClosePosit:=Pos(C2,S[I]);
      If (Posit<>0) and (ClosePosit>Posit) then
      repeat
        Buf:=S[I];
        NotesBody.Add('<section id="FbAutId_'+IntToStr(ID)+'"><title><p>Note'+
          IntToStr(ID)+'</p></title>');
        NotesBody.Add('<p>'+CleanTags(Copy(Buf,Posit+Length(C1),ClosePosit-Posit-Length(C1)))+'</p></section>');
        Delete(Buf,Posit,ClosePosit-Posit+Length(c2));
        Insert('<a xlink:href="#FbAutId_'+IntToStr(ID)+
        '" type="note">note '+IntToStr(ID)+'</a>',Buf,Posit);
        S[I]:=Buf;     
        Inc(ID);
        Posit:=Pos(C1,S[I]);
        ClosePosit:=Pos(C2,S[I]);
      until (Posit=0) or (ClosePosit<=Posit);
    end;
  Begin
    If Params.NotDetectNotes then Exit;
    Inform('Detecting notes...');
    I:=0;
    ID:=1;
    WHile I< S.Count-2 do
    Begin
      CheckBlock('[',']');
      CheckBlock('{','}');
      Inc(I);
    end;
{    if NotesBody.Count<>0 then
      NotesBody.Add('</body>')}
  end;
  Function PostProcess(const S:String):String;
  Begin
    inform('Post-processing text...');
    RegExp.ReplaceOnRequest(S,'(<p><(a|strong|emphasis)[^>]*>)\s+','$1',True);
    Progress(10);
    RegExp.ReplaceOnRequest('','</section>[^<]*<section><title><p>([^<]*)</p></title>[^<]*</section>','<subtitle>$1</subtitle></section>',True);
    Progress(20);
    RegExp.ReplaceOnRequest('','<(?=[^>]*<)','&lt;',True);
    Progress(30);
    RegExp.ReplaceOnRequest('','<empty-line/>[\s'#10#13']*</section>','</section>',True);
    Progress(35);
    RegExp.ReplaceOnRequest('','<(emphasis|strong)>(([^<]*)<\1>)+','$3<$1>',True);
    Progress(40);
    Result:=RegExp.ReplaceOnRequest('','(>[^<]*?)>','$1&gt;',True);
    Progress(50);
    Result:=RegExp.ReplaceOnRequest('','<p id="([^"]*)">','<p id="'+IDPrefix+'$1">',True);
    Progress(60);
    Result:=RegExp.ReplaceOnRequest('','<a xlink:href="#([^"]*)">','<a xlink:href="#'+IDPrefix+'$1">',True);
    Progress(90);
    ReplaceStr(Result,#3,'');
    Progress(100);
  end;


Begin
  InputLines:=ValidHtml(InputLines);
  FormatQ(InputLines);
  RemoveFormsScripts(InputLines);
  DetectParagraphs(InputLines);
  DetectHeaders(InputLines);
  DetectEpigraph(InputLines);
  CreateParagraphs(InputLines);
  ItalicCreate(InputLines);
  RemoveAllDSpaces(InputLines);
  KillEntity(InputLines);
  FormatT(InputLines);
  FormatP(InputLines);
  CreateFootnotes(InputLines);
  RemoveNotClosed(InputLines);
  RemoveAllDSpaces(InputLines);
  DetectVerses(InputLines);
  DetectNesting(InputLines);
  InputLines.Text:=PostProcess(InputLines.Text);
end;
//==============================================================================
//==============================================================================


Function ValidateText(LinesInWork:TStringList;FN:String):Boolean;
Var
  XDoc:IXMLDOMDocument2;
  FixTrialsDone,PrevError,PrePrevError:Integer;
  RegExpList:TStringList;
  I:Integer;
  ConvertResult:WideString;
Begin
  if Done<>0 then
    LinesInWork.Insert(0,'<body name="'+MyExtractFileName(FN)+'" xmlns:fb="http://www.gribuser.ru/xml/fictionbook/2.0" xmlns:xlink="http://www.w3.org/1999/xlink">')
  else
    LinesInWork.Insert(0,'<body xmlns:fb="http://www.gribuser.ru/xml/fictionbook/2.0" xmlns:xlink="http://www.w3.org/1999/xlink">');
  LinesInWork.Add('</body>');
//  LinesInWork.SaveToFile('c:\temp\temp.xml');
{      LinesInWork.Text:=ConvertResult;
    LinesInWork.Free;}
  FixTrialsDone:=-50;
  PrevError:=0;
  PrePrevError:=0;
  try
    XDoc:=CoDOMDocument40.Create;
  except
    raise exception.Create('MSXML4.0 init failed!')
  end;
  Result:=False;
  if DirectoryExists('c:\any2fbtemp') then
    LinesInWork.SaveToFile('c:\any2fbtemp\before_check.xml');
  Repeat
    XDoc.loadXML(LinesInWork.Text);
    if XDoc.parseError.errorCode<>0 then
    Begin
      Inc(FixTrialsDone);
      If XDoc.parseError.line<2 then Break;
      If XDoc.parseError.line<>PrevError then
      Begin
        Warn('XML Validation failed ('+IntToStr(FixTrialsDone+50)+' pass) at line '+intToStr(XDoc.parseError.line)+', in text: /'+
          CleanTags(LinesInWork[XDoc.parseError.line-1])+'/'#13'Fixing and retrying...');
        if pos('<section><title>',LinesInWork[XDoc.parseError.line-1])=0 then
          LinesInWork[XDoc.parseError.line-1]:='<p><style name="converterror">'+CleanTags(LinesInWork[XDoc.parseError.line-2])+'</style></p>'
        else
          LinesInWork[XDoc.parseError.line-1]:='<section><title><p><style name="converterror">'+CleanTags(LinesInWork[XDoc.parseError.line-2])+'</style></p></title>';
      end else
        if PrePrevError<>PrevError then
        Begin
          LinesInWork[XDoc.parseError.line-2]:='<p><style name="converterror">'+CleanTags(LinesInWork[XDoc.parseError.line-3])+'</style></p>';
          LinesInWork[XDoc.parseError.line-1]:='<p><style name="converterror">'+CleanTags(LinesInWork[XDoc.parseError.line-2])+'</style></p>';
          LinesInWork[XDoc.parseError.line]:='<p><style name="converterror">'+CleanTags(LinesInWork[XDoc.parseError.line-1])+'</style></p>';
          Warn('XML Validation failed TWICE at line '+intToStr(XDoc.parseError.line)+', fixing hard (3 lines) and retrying...');
          PrePrevError:=XDoc.parseError.line;
        end else
          Begin
            Warn('XML Validation failed THREE TIMES at line '+intToStr(XDoc.parseError.line)+', completly removing damaged line and retrying.'#10+
              'Removed text is:'#10+LinesInWork[XDoc.parseError.line-2]+#10);
            LinesInWork[XDoc.parseError.line-1]:='<p><style name="converterror">In this place was an unrecoverable import error. Sorce text was completly removed, so you may need to re-enter it here manually</style></p>';
            PrePrevError:=0;
          end;
      PrevError:=XDoc.parseError.line;
    end;
  until (XDoc.parseError.errorCode=0) or (FixTrialsDone>=Params.FixTrialsOver100);

  If XDoc.parseError.errorCode=0 then
  Begin
    Inform('XML validation passed OK');
    Result:=True;
  end
  else
    Begin
      Warn('Unable to fix invalid XML generated. Aborting convertion!');
      if DirectoryExists('c:\any2fbtemp') then
        LinesInWork.SaveToFile('c:\any2fbtemp\broken.xml')
      else
        Warn('Create folder c:\any2fbtemp to view invalid document.');
      Exit;
    end;
  if (Params.RegExpOnFinish<>'') then
  Begin
    Inform('Running user regular expressions on ready document...');
    RegExp.InputText:=LinesInWork.Text;
    RegExpList:=TStringList.Create;
    RegExpList.Text:=Params.RegExpOnFinish;
    If RegExpList.Count<=((RegExpList.Count-1) div 2)*2+1 then
      RegExpList.Add('');
    For I:=0 to (RegExpList.Count-1) div 2 do
    Begin
      Inform(#9'/'+RegExpList[I*2]+'/'+RegExpList[I*2+1]+'/');
      ConvertResult:=RegExp.ReplaceOnRequest('',RegExpList[I*2],RegExpList[I*2+1],True);
    end;
    RegExpList.Free;
    XDoc.loadXML(ConvertResult);
    If XDoc.parseError.errorCode<>0 then
    Begin
      Result:=False;
      Warn('XML is invalid after running user Regular Expressions');
      if DirectoryExists('c:\any2fbtemp') then
        LinesInWork.SaveToFile('c:\any2fbtemp\broken.xml')
      else
        Warn('Create folder c:\any2fbtemp to view invalid document.');
      Exit;
    end;
    LinesInWork.Text:=ConvertResult;
  end;
end;

Function BuildOneBody(AText:String;Params:PParceParam;FootNotes,UsedFiles,BInaryes:TStringList):String;
Var
  LinesInWork:TStringList;
  HRBase,RootPath,Buf:String;
  LevelSeparator:Char;
  I:Integer;

  Procedure FindBaseHRef(S:String);
  Begin
    HRBase:=S;
    If pos('http://',LowerCase(S))=1 then
    Begin
      LevelSeparator:='/';
      SetLength(S,Length(S)-Length(StrRScan(PChar(S),'/'))+1);
      If UpperCase(S)<>'HTTP://' then
      Begin
        HRBase:=S;
        RootPath:=Copy(S,1,Pos('/',Copy(S,8,MaxInt))+6);
      end else RootPath:=HRBase;
      If RootPath[Length(RootPath)]='/' then
        SetLength(RootPath,Length(RootPath)-1);
      If HRBase[Length(HRBase)]='/' then
        SetLength(HRBase,Length(HRBase)-1);
    end else
      Begin
        LevelSeparator:='\';
        SetLength(HRBase,Length(HRBase)-Length(StrRScan(PChar(HRBase),'\'))+1);
        RootPath:=Copy(HRBase,1,3);
      end;
    RootPath:=LowerCase(RootPath);
  end;
  Function ExpandHRef(HRef:String):String;
  Var
    NewHRef:String;
  Begin
    Result:=FTrim(HRef);
    If Result='' then Exit;
    If Pos('mailto:',LowerCase(HRef))=1 then
    Begin
      Result:=HRef;
      Exit;
    end;
    If (pos('http://',LowerCase(HRef))=1) or (Pos(':',HRef)=2) then Exit;
    If Result[1]='/' then
    Begin
      Result:=RootPath+Result;
      Exit;
    end;
    NewHRef:=HRBase;
    While pos('../',Result)=1 do
    Begin
      SetLength(NewHRef,Length(NewHRef)-Length(StrRScan(PChar(Copy(NewHRef,1,Length(NewHREf)-1)),LevelSeparator)));
      Result:=Copy(Result,4,MaxInt);
    end;
    If NewHRef[Length(NewHRef)]<>LevelSeparator then
      Result:=NewHRef+LevelSeparator+Result
    else
      Result:=NewHRef+Result;
    If LevelSeparator='\' then
      ReplaceStr(Result,'/','\');
  end;
  Procedure SplitHRef(Var HR,ID:String);
  Begin
    If Pos('mailto:',LowerCase(HR))=1 then
    Begin
      HR:='';
      ID:='';
      Exit;
    end;
    If Pos('#',HR)=0 then
    Begin
      id:='';
      Exit;
    end;
    ID:=Copy(HR,Pos('#',HR)+1,MaxInt);
    SetLength(HR,Pos('#',HR)-1);
  end;
  Procedure CollectHRefs(S:TStringList;MayDeeper:Boolean);
  Var
    I,I1:Integer;
    Buf,Substr,FoundHR,FoundID,RE,REP:String;
  Begin
    If params.RemoveExternalLinks then Exit;
    Inform('Collecting external links...');
    For I:=0 to S.Count-1 do
      If Pos('<a xl',S[I])<>0 then
        Begin
          Progress(Round((I/S.Count)*100));
          Buf:=S[I];
          While pos('<a xlink:href="',Buf)<>0 do
            Begin
              Buf:=Copy(Buf,pos('<a xlink:href="',Buf)+15,MaxInt);
              FoundHR:=Copy(Buf,1,Pos('"',Buf)-1);

              SplitHRef(FoundHR,FoundID);
              Substr:=ExpandHRef(FoundHR);
              If (FoundHR<>'') and ((Pos(RootPath,LowerCase(Substr))=1) or (Params.DownloadExternal))
              and (RegExp.Match(Substr,Params.LinksFollow) or
              (Params.LinksFollow='')) and ((Params.LinksSkip='') or not RegExp.Match(Substr,Params.LinksSkip)) then
              Begin
                if MayDeeper or (UsedFiles.IndexOf(Substr)<>-1) then
                Begin
                  RE:='<a xlink:href="'+FoundHR;
                  If FoundID<>'' then RE:=RE+'#'+FoundID;
                  RE:=RE+'"';
                  FoundHR:=ExpandHRef(FoundHR);
                  ReplaceStr(re,'\','\\');
                  ReplaceStr(RE,'.','\.');
                  ReplaceStr(RE,'?','\?');
                  ReplaceStr(RE,'+','\+');
                  ReplaceStr(RE,'*','\*');
                  ReplaceStr(RE,'(','\(');
                  ReplaceStr(RE,')','\)');
                  I1:=UsedFiles.IndexOf(FoundHR);
                  If I1<0 then
                    I1:=UsedFiles.Add(FoundHR);
                  REp:='<a xlink:href="#AutBody_'+IntToStr(I1);
                  If FoundID<>'' then
                  Begin
                    If FoundID[1] in ['0'..'9'] then
                      FoundID:='fb_'+FoundID;
                    ReplaceStr(FoundID,'"','_');
                    Rep:=Rep+FoundID
                  end else Rep:=Rep+'DocRoot';
                  Rep:=Rep+'"';
                  S[I]:=RegExp.ReplaceOnRequest(S[I],RE,REP,False);
                end
                else
              end;
            end;
        end;
  end;
  Procedure MarkStart(S:TStringList);
  Var
    I:Integer;
  Begin
    For I:=0 to S.Count-1 do
      If Pos('<p>',S[I])=1 then
      Begin
        S[I]:='<p id="'+IDPrefix+'DocRoot"'+Copy(S[I],3,MaxInt);
        Exit;
      end;
  end;
  Procedure LoadImages;
  Var
   I,ImgPos:Integer;
   Buf1,Buf2,TAG,URL:String;
   Function IsDinamic(URL:String):Boolean;
   Begin
     URL:=LowerCase(URL);
     Result:=(Pos('?',URL)<>0) or (pos('.php',URL)<>0) or (pos('.cgi',URL)<>0
     ) or (pos('.pl',url)<>0) or (pos('.asp',url)<>0) or (pos('.jsp',url)<>0)
     or (pos('banner',url)<>0)or (pos('adv',url)<>0);
   end;
  Begin
    For I:=0 to LinesInWork.Count-1 do
      Begin
        ImgPos:=pos('<imagex',LinesInWork[I]);
        While ImgPos<>0 do
          Begin
            Buf1:=Copy(LinesInWork[I],1,ImgPos-1);
            TAG:=Copy(LinesInWork[I],ImgPos,MaxInt);
            Tag:=Copy(Tag,1,Pos('>',Tag));
            Buf2:=Copy(LinesInWork[I],Length(Buf1)+Length(Tag)+1,MaxInt);
            URL:=Copy(TAG,Pos('"',Tag)+1,MaxInt);
            URL:=Copy(URL,1,Pos('"',URL)-1);
            URL:=ExpandHRef(URL);
            if ((Pos(RootPath,LowerCase(URL))=1) or not Params.RemoveExternalLinks)
              and (not IsDinamic(URL) or Params.KeepDynamicImages) then
              Begin
                If BInaryes.IndexOf(URL)=-1 then
                  BInaryes.Add(URL);
                Tag:='<image xlink:href="#Any2FbImgLoader'+IntToStr(BInaryes.IndexOf(URL))+'"/>';
              end
            else
              TAG:='';
            If (Buf1='<p>') and (Buf2='</p>') then
              LinesInWork[I]:=TAG
            else
              LinesInWork[I]:=Buf1+TAG+Buf2;
            ImgPos:=pos('<imagex',LinesInWork[I]);
          end;
      end;
  end;
Begin
  Result:='';

  Description:='';
  LinesInWork:=GetTheText(AText,Params);
  If UsedFiles.IndexOf(AText)<0 then
    UsedFiles.Add(AText);
  IDPrefix:='AutBody_'+IntToStr(UsedFiles.IndexOf(AText));
  RecogniseText(LinesInWork,FootNotes);
  Try
    If not ValidateText(LinesInWork,AText) then
      Exit;
    MarkStart(LinesInWork);
    UsedFiles.Objects[UsedFiles.IndexOf(AText)]:=sReady;
    FindBaseHRef(AText);
    If not Params.SkipAllImages then
      LoadImages;

    Result:=LinesInWork.Text;
    If Params.LinksDeep>0 then
    Begin
      CollectHRefs(LinesInWork,True);
      Result:=LinesInWork.Text;
      Inform('Going to download linked from "'+AText+'" files...');
      I:=0;
      Dec(Params.LinksDeep);
      While I<UsedFiles.Count do
      Begin
        If UsedFiles.Objects[I]=Nil then
        Try
          Inform('('+IntToStr(I+1)+'/'+IntToStr(UsedFiles.Count)+') Working with linked URL '+UsedFiles[I]);
          Buf:=UsedFiles[I];
          Result:=Result+BuildOneBody(Buf,Params,FootNotes,UsedFiles,BInaryes);
          UsedFiles[I]:=Buf;
          Inc(Done);
        except
          Warn('ERROR working with '+UsedFiles[I]+'!!! - skipped');
          Inc(Skipped);
        end;
        Inc(I);
      end;
      Inc(Params.LinksDeep);
    end else
      Begin
        CollectHRefs(LinesInWork,False);
        Result:=LinesInWork.Text;
      end;
  Finally
    LinesInWork.Free;
  end;
end;
Procedure DownloadImages(ImgList:TStringList);
Var
  I:Integer;
  UUECOmpressor:TNMUUProcessor;
  HTTPSocket:TNMHTTP;
  Buf:String;
  ResultSize:Integer;
  ContentType:String;
  Ext:String;
begin
  If (ImgList.Count=0) then Exit;
  Inform('Loading images ('+IntToStr(ImgList.Count)+')...');
  UUECOmpressor:=TNMUUProcessor.Create(Nil);
  Try
    UUECOmpressor.InputStream:=Nil;
    UUECOmpressor.Method:=uuMime;
    UUECOmpressor.OutputStream:=TMemoryStream.Create;
    For I:=0 to ImgList.Count-1 do
      Try
        Buf:=ImgList[I];
        ReplaceStr(Buf,'%20',' ');
        ImgList[I]:=Buf;
        Inform('('+IntToStr(I+1)+'/'+IntToStr(ImgList.Count)+') '+ImgList[I]);
        Try
          UUECOmpressor.OutputStream.Size:=0;
          If Pos('http://',LowerCase(ImgList[I]))=1 then
          Begin
            HTTPSocket:=TNMHTTP.Create(Nil);
            Try
              HTTPSocket.OnPacketRecvd:=TNotifyEvent(ProgressFunction);
              Try
                UUECOmpressor.InputStream:=Nil;
                HTTPSocket.Get(ImgList[I]);
                UUECOmpressor.InputStream:=TMemoryStream.Create;
                Buf:=HTTPSocket.Body;
                UUECOmpressor.InputStream.Write(Buf[1],Length(Buf));
                UUECOmpressor.InputStream.Seek(0,soFromBeginning);
                If pos('image/jpeg',lowercase(HTTPSocket.Header))<>0 then
                  ContentType:='image/jpeg'
                else If pos('image/png',lowercase(HTTPSocket.Header))<>0 then
                    ContentType:='image/png'
                else If pos('image/gif',lowercase(HTTPSocket.Header))<>0 then
                    ContentType:='image/gif'
                {else If (pos('.jpg',lowercase(ImgList[I]))<>0) or (pos('.jpeg',lowercase(ImgList[I]))<>0) then
                    ContentType:='image/jpeg'
                else If pos('.png',lowercase(ImgList[I]))<>0 then
                    ContentType:='image/png'
                else If pos('.gif',lowercase(ImgList[I]))<>0 then
                    ContentType:='image/gif'}
                  else ContentType:='image';
              except
                ContentType:='image';
                Warn('Error downloading image!');
                UUECOmpressor.InputStream.Free;
                UUECOmpressor.InputStream:=Nil;
              end
            Finally
              HTTPSocket.Free;
            end;
          end else
            Begin
              Try
                UUECOmpressor.InputStream:=TFileStream.Create(ImgList[I],fmOpenRead);
                Buf:=ExtractFileName(ImgList[I]);
                If (Pos('.jpg',Buf)<>0) or (Pos('.jpeg',Buf)<>0) then
                  ContentType:='image/jpeg'
                else If (Pos('.png',Buf)<>0)then
                  ContentType:='image/png'
                else If (Pos('.gif',Buf)<>0)then
                  ContentType:='image/gif'
                else ContentType:='image';
              Except
                UUECOmpressor.InputStream:=Nil;
                Warn('Not loaded!');
              end;
            end;
          if ContentType='image/gif' then
          Begin
            Try
              Inform(#9'Converting gif->png');
              UUECOmpressor.InputStream:=PNGFromGif(UUECOmpressor.InputStream);
            except
            end;
            ContentType:='image/png';
            UUECOmpressor.InputStream:=PNGSteam;
          end;
          UUECOmpressor.Encode;
        Finally
          UUECOmpressor.InputStream.Free;
          UUECOmpressor.InputStream:=Nil;
        end;
        Ext:='jpg';
        if ContentType='image/png' then
        Begin
          Ext:='png';
          ImgList.Objects[I]:=pointer(1);
        end;
        ResultSize:=UUECOmpressor.OutputStream.Seek(0,soFromEnd);
        UUECOmpressor.OutputStream.Seek(0,soFromBeginning);
        SetLength(Buf,ResultSize);
        UUECOmpressor.OutputStream.Read(Buf[1],ResultSize);
        ImgList[I]:='<binary content-type="'+ContentType+'" id="Any2FbImgLoader'+IntToStr(I)+'.'+Ext+'">'+
          Buf+'</binary>';

        If ContentType='image' then Warn(#9'Invalid image format!')
        else Inform(#9'OK');
      Except
        ImgList[I]:='';
      end;
  finally
    UUECOmpressor.OutputStream.Free;
    UUECOmpressor.Free;
  end;
end;
Procedure CleanFolder;
Var
  FolderName:String;
  sr: TSearchRec;
  FileList:TStringList;
  I:Integer;
Begin
  SysUtils.DeleteFile(NedClean);

  FolderName:=Copy(NedClean,1,Length(NedClean)-Length(ExtractFileExt(NedClean)))+'.files';

  if SysUtils.FindFirst(FolderName+'\*.*', faAnyFile , sr) = 0 then
  begin
    FileList:=TStringList.Create;
    Try
      repeat
        if (sr.Name<>'..') and (sr.Name<>'.') then
          FileList.Add(sr.Name);
      until FindNext(sr) <> 0;
      SysUtils.FindClose(sr);
      For I:=0 to FileList.Count-1 do
        SysUtils.DeleteFile(FolderName+'\'+FileList[I]);
    finally
      FileList.Free;
    end;
  end;
  RemoveDirectory(PChar(FolderName));
end;

Function CollectHead:String;
Begin
  Result:='<?xml version="1.0" encoding="Windows-1251"?>'+
  '<FictionBook xmlns:xlink="http://www.w3.org/1999/xlink" xmlns="http://www.gribuser.ru/xml/fictionbook/2.0">'+
  '<description><title-info><genre></genre><author><first-name>'+
  DelEnt(FAutor)+'</first-name><middle-name>'+DelEnt(MAuthor)+
  '</middle-name><last-name>'+DelEnt(LAuthor)+'</last-name></author>'+
   '<book-title>'+DelEnt(Title)+'</book-title>';
  If Description<>'' then
    Result:=Result+'<annotation>'+DelEnt(Description)+'</annotation>';
  Result:=Result+'</title-info></description>';
end;

begin
  InformFunction:=AInformFunction;
  WarningFunction:=AWarningFunction;
  ProgressFunction:=ACallBackFunction;
  FootNotes:=TStringList.Create;
  UsedFiles:=TStringList.Create;
  BInaryes:=TStringList.Create;
  Try
    Try
      if Params=Nil then
      Begin
        New(Params);
        FillChar(Params^,SizeOf(Params^),0);
        Params.ConvertLong:=2;
      end;
      Done:=0;
      Skipped:=0;

      DonePart:=BuildOneBody(AText,Params,FootNotes,UsedFiles,BInaryes);
      If DonePart='' then
      Begin
        Result:=Nil;
        Exit;
      end;
      Try
        DownloadImages(BInaryes);
      except
        Warn('Error downloading images...');
      end;

      XMLHead:=CollectHead;
      XDoc:=CoFreeThreadedDOMDocument40.create;
      If FootNotes.Count>0 then
        XDoc.loadXML(XMLHead+DonePart+'<body name="notes">'+FootNotes.text+'</body>'+BInaryes.Text+'</FictionBook>')
      else
        XDoc.loadXML(XMLHead+DonePart+BInaryes.Text+'</FictionBook>');
      If XDoc.parseError.errorCode<>0 then
      Begin
        Description:='<p>'+CleanTags(Description)+'</p>';
        XMLHead:=CollectHead;
        If FootNotes.Count>0 then
          XDoc.loadXML(XMLHead+DonePart+'<body name="notes">'+FootNotes.text+'</body>'+BInaryes.Text+'</FictionBook>')
        else
          XDoc.loadXML(XMLHead+DonePart+BInaryes.Text+'</FictionBook>');
        Result:=Nil;
        If XDoc.parseError.errorCode<>0 then
          Raise Exception.Create('Error adding document header.'#10+
            'Try editing <head> section in html document and turn off description generation');
      end;
      Inform('Total linked documents imported - '+IntToStr(Done));
      Inform('Total linked documents skipped - '+IntToStr(Skipped));
      WholeDoc:=XDoc.xml;
      if BInaryes.Count>0 then
      begin
        for I:= 0 to BInaryes.Count-1 do
        begin
          if BInaryes.Objects[I] = nil then
            Ext:='jpg'
          else
            Ext:='png';
          WholeDoc:=RegExp.ReplaceOnRequest(WholeDoc,'(#Any2FbImgLoader'+IntToStr(I)+')','$1.'+Ext,false);
        end;
        WholeDoc:=RegExp.ReplaceOnRequest(WholeDoc,'<\?xml version="1\.0"\?>','<?xml version="1.0" encoding="UTF-16"?>',false);
        XDoc.loadXML(WholeDoc);
      end;
      Result:=XDoc;
    except
      On E:exception do Begin
        Warn('Internal error: '+E.Message);
      end;
    end;
  finally
    if NedClean<>'' then
      CleanFolder;
    FootNotes.Free;
    UsedFiles.Free;
    BInaryes.Free;
  end;
end;
begin
  RegExp:=TMSRegExpr.Create;
  RegExp.CaseSencitive:=False;
end.
