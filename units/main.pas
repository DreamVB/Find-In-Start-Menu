unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  LCLIntf, Buttons, LCLType, ComCtrls, Menus, Windows, ActiveX, ShlObj;

type

  { Tfrmmain }

  Tfrmmain = class(TForm)
    cmdOpenPath: TSpeedButton;
    cboStartLoc: TComboBox;
    ImgIcons: TImageList;
    lblStartMnuPaths: TLabel;
    lblFound: TLabel;
    lblSearch: TLabeledEdit;
    lstFiles: TListBox;
    cmdRun: TSpeedButton;
    mnuOpenFolder: TMenuItem;
    mnuRun: TMenuItem;
    mnuLst: TPopupMenu;
    StatusBar1: TStatusBar;
    procedure Button1Click(Sender: TObject);
    procedure cboStartLocChange(Sender: TObject);
    procedure cmdOpenPathClick(Sender: TObject);
    procedure cmdRunClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure lblSearchChange(Sender: TObject);
    procedure lstFilesDblClick(Sender: TObject);
    procedure lstFilesDrawItem(Control: TWinControl; Index: integer;
      ARect: TRect; State: TOwnerDrawState);
    procedure lstFilesKeyPress(Sender: TObject; var Key: char);
    procedure lstFilesMeasureItem(Control: TWinControl; Index: integer;
      var AHeight: integer);
    procedure mnuOpenFolderClick(Sender: TObject);
    procedure mnuRunClick(Sender: TObject);
    procedure mnuLstPopup(Sender: TObject);
  private
    function FixPath(S: string): string;
    procedure AddProgramsToListBox(sl: TStringList; const StartDir: string);
    procedure SearchForApp(Source: TStringList; Dest: TStringList; sFind: string);
    function GetFileTitle(S: string): string;
    procedure LoadPaths(Loc: string);
  public

  end;

var
  frmmain: Tfrmmain;
  TPrograms: TStringList;
  TFindFiles: TStringList;
  TFindInex: TStringList;

implementation

{$R *.lfm}

{ Tfrmmain }

function GetUserDir: string;
var
  nSize: DWORD;
begin
  nSize := 1024;

  SetLength(Result, nSize);
  if GetUserName(PChar(Result), nSize) then
  begin
    SetLength(Result, nSize - 1);
  end;
end;

function GetShortcutTarget(const ShortcutFilename: string): string;
var
  Psl: IShellLink;
  Ppf: IPersistFile;
  WideName: array[0..MAX_PATH] of widechar;
  pResult: array[0..MAX_PATH - 1] of ansichar;
  Data: TWin32FindData;
const
  IID_IPersistFile: TGUID = (
    D1: $0000010B; D2: $0000; D3: $0000; D4: ($C0, $00, $00, $00, $00, $00, $00, $46));
begin
  CoCreateInstance(CLSID_ShellLink, nil, CLSCTX_INPROC_SERVER, IID_IShellLinkA, psl);
  psl.QueryInterface(IID_IPersistFile, ppf);
  MultiByteToWideChar(CP_ACP, 0, pansichar(ShortcutFilename), -1, WideName, Max_Path);
  ppf.Load(WideName, STGM_READ);
  psl.Resolve(0, SLR_ANY_MATCH);
  psl.GetPath(@pResult, MAX_PATH, Data, SLGP_UNCPRIORITY);
  Result := StrPas(pResult);
end;

procedure Tfrmmain.LoadPaths(Loc: string);
begin
  TPrograms := TStringList.Create;
  TFindInex := TStringList.Create;
  TFindFiles := TStringList.Create;
  //Get a list of all start menu items
  AddProgramsToListBox(TPrograms, Loc);
end;

function Tfrmmain.GetFileTitle(S: string): string;
var
  sPos: integer;
  lzFile: string;
begin
  lzFile := ExtractFileName(S);
  //Get pos of .
  sPos := Pos('.', lzFile);
  //Check if found
  if sPos > 0 then
  begin
    Result := LeftStr(lzFile, sPos - 1);
  end
  else
  begin
    Result := S;
  end;
end;

procedure Tfrmmain.AddProgramsToListBox(sl: TStringList; const StartDir: string);
var
  sr: TSearchRec;
  lzFile, fExt: string;
begin

  if FindFirst(FixPath(StartDir) + '*.*', faAnyFile or faDirectory, sr) = 0 then
  begin
    repeat
      if (sr.Attr and faDirectory) = 0 then
      begin
        lzFile := FixPath(StartDir) + sr.Name;
        fExt := LowerCase(ExtractFileExt(lzFile));
        if fExt = '.lnk' then
        begin
          sl.Add(lzFile);
        end;
      end
      else if (sr.Name <> '.') and (sr.Name <> '..') then
      begin
        AddProgramsToListBox(sl, FixPath(StartDir) + sr.Name);
      end;
    until FindNext(sr) <> 0;
  end;

end;

procedure Tfrmmain.SearchForApp(Source: TStringList; Dest: TStringList; sFind: string);
var
  X: integer;
  lzFile, lzCheckFile: string;
begin
  for X := 0 to Source.Count - 1 do
  begin
    lzFile := Source[X];
    lzCheckFile := Lowercase(GetFileTitle(lzFile));
    if Pos(lowercase(sFind), lzCheckFile) <> 0 then
    begin
      Dest.Add(GetFileTitle(lzFile));
      TFindFiles.Add(lzFile);
    end;
  end;
end;

function Tfrmmain.FixPath(S: string): string;
begin
  if rightstr(S, 1) <> PathDelim then
  begin
    Result := S + PathDelim;
  end
  else
  begin
    Result := S;
  end;
end;

procedure Tfrmmain.cmdRunClick(Sender: TObject);
begin
  lstFilesDblClick(Sender);
end;

procedure Tfrmmain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  FreeAndNil(TPrograms);
  FreeAndNil(TFindFiles);
  FreeAndNil(TFindInex);
end;

procedure Tfrmmain.cmdOpenPathClick(Sender: TObject);
var
  idx: integer;
  lzPath: string;
begin
  idx := lstFiles.ItemIndex;

  if idx <> -1 then
  begin
    lzPath := ExtractFilePath(GetShortcutTarget(TFindFiles[idx]));
    OpenDocument(lzPath);
  end;

end;

procedure Tfrmmain.cboStartLocChange(Sender: TObject);
var
  idx: integer;
begin
  idx := cboStartLoc.ItemIndex;
  //Clear files list first
  lstFiles.Clear;
  if idx <> -1 then
  begin
    LoadPaths(cboStartLoc.Items[idx]);
    lblSearchChange(Sender);
  end;
end;

procedure Tfrmmain.Button1Click(Sender: TObject);
begin

end;

procedure Tfrmmain.FormCreate(Sender: TObject);
begin

  cboStartLoc.Items.Add('C:\ProgramData\Microsoft\Windows\Start Menu\');
  cboStartLoc.Items.Add('C:\Users\' + GetUserDir +
    '\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\');
  cboStartLoc.Items.Add('C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp');

  cboStartLoc.ItemIndex := 0;
  cboStartLocChange(Sender);

end;

procedure Tfrmmain.FormPaint(Sender: TObject);
begin

end;

procedure Tfrmmain.lblSearchChange(Sender: TObject);
begin

  if TFindInex <> nil then
  begin
    TFindInex.Clear;
    TFindFiles.Clear;
  end;

  SearchForApp(TPrograms, TFindInex, lblSearch.Text);

  if (TFindInex <> nil) then
  begin
    lblFound.Caption := 'Found: ' + IntToStr(TFindInex.Count) + ' items(s)';
    lstFiles.Items.Assign(TFindInex);
  end;
end;

procedure Tfrmmain.lstFilesDblClick(Sender: TObject);
var
  idx: integer;
  lzAppFile: string;
begin

  idx := lstFiles.ItemIndex;

  if (idx <> -1) then
  begin
    //Get shortct target
    lzAppFile := GetShortcutTarget(TFindFiles[idx]);

    if FileExists(lzAppFile) then
    begin
      if not OpenDocument(lzAppFile) then
      begin
        MessageDlg(Text, 'There was an error executing the application:' +
          sLineBreak + sLineBreak + lzAppFile, mtError, [mbOK], 0);
      end;
    end
    else
    begin
      MessageDlg(Text, 'File Was Not Found:' + sLineBreak + sLineBreak + lzAppFile,
        mtWarning, [mbOK], 0);
    end;
  end;

end;

procedure Tfrmmain.lstFilesDrawItem(Control: TWinControl; Index: integer;
  ARect: TRect; State: TOwnerDrawState);
var
  YPos: integer;
begin
  if odSelected in State then
  begin
    lstFiles.Canvas.Brush.Color := $00BE5FA9;
  end;

  //Draw the icons in the listbox
  lstFiles.Canvas.FillRect(ARect);
  //Draw the icon on the listbox from the image list.
  ImgIcons.Draw(lstFiles.Canvas, ARect.Left + 4, ARect.Top + 4, 0);
  //Align text
  YPos := (ARect.Bottom - ARect.Top - lstFiles.Canvas.TextHeight(Text)) div 2;

  //Write the list item text
  lstFiles.Canvas.TextOut(ARect.left + ImgIcons.Width + 8, ARect.Top + YPos,
    lstFiles.Items.Strings[index]);
end;

procedure Tfrmmain.lstFilesKeyPress(Sender: TObject; var Key: char);
begin
  if Key = #13 then
  begin
    Key := #0;
    lstFilesDblClick(Sender);
  end;
end;

procedure Tfrmmain.lstFilesMeasureItem(Control: TWinControl; Index: integer;
  var AHeight: integer);
begin
  AHeight := ImgIcons.Height + 8;
end;

procedure Tfrmmain.mnuOpenFolderClick(Sender: TObject);
begin
  cmdOpenPathClick(Sender);
end;

procedure Tfrmmain.mnuRunClick(Sender: TObject);
begin
  cmdRunClick(Sender);
end;

procedure Tfrmmain.mnuLstPopup(Sender: TObject);
var
  idx: integer;
begin
  idx := lstFiles.ItemIndex;
  mnuRun.Enabled := (idx <> -1);
  mnuOpenFolder.Enabled := mnuRun.Enabled;
end;

end.
