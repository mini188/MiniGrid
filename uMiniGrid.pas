{***************************************************************************}
{ TMiniGrid component                                                       }
{ for Delphi                                                                }
{                                                                           }
{ written by mini188                                                        }
{ Email : mini188.com@qq.com                                                }
{                                                                           }
{***************************************************************************}

unit uMiniGrid;

interface
uses
  Windows, Classes, Contnrs, Grids, SysUtils, RTLConsts, Graphics, Controls,
  Messages, ShellAPI, StrUtils, IniFiles;

type
  TMiniGrid = class;
  TColumns = class;

  TUrlCache = class
    FullData: string;
    Url: string;
    Text: string;
    Col: Integer;
    Row: Integer;
  end;

  TColumn = class(TComponent)
  private
    FOwner: TColumns;
    FTitle: string;
    FFieldName: string;
    FWidth: Integer;
    FFontColor: TColor;
    FHeadBackgroundColor: TColor;
    FFont: TFont;
    procedure SetFieldName(const Value: string);
    procedure SetTitle(const Value: string);
    procedure SetWidth(const Value: Integer);
  public
    constructor Create(AOwner: TColumns);
    destructor Destroy; override;
    property Title: string read FTitle write SetTitle;
    property FieldName: string read FFieldName write SetFieldName;
    property Width: Integer read FWidth write SetWidth;
    property Font: TFont read FFont write FFont;
    property HeadBackgroundColor: TColor read FHeadBackgroundColor write FHeadBackgroundColor; 
  end;

  TColumns = class(TObjectList)
  private
    FOwnerGrid: TMiniGrid;
    function GetColumn(Index: Integer): TColumn;
  public
    constructor Create(AOwner: TMiniGrid);
    destructor Destroy; override;

    function AddColumn: TColumn;
    procedure DeleteColumn(AIndex: Integer);
    procedure Changed;

    property Columns[Index: Integer]: TColumn read GetColumn;
  end;

  TMergeInfo = class(TPersistent)
  private
    FRowSpan: Integer;
    FRow: Integer;
    FColSpan: Integer;
    FCol: Integer;
    FPaintId: Integer;
  public
    property Col: Integer read FCol write FCol;
    property Row: Integer read FRow write FRow;
    property RowSpan: Integer read FRowSpan write FRowSpan;
    property ColSpan: Integer read FColSpan write FColSpan;
    property PaintId: Integer read FPaintId write FPaintId;
  end;

  TMergeInfos = class(TObjectList)
  private
    function GetMergeInfo(AIndex: Integer): TMergeInfo;
  public
    function IsMergeCell(ACol, ARow: Integer): Boolean;
    function IsBaseCell(ACol, ARow: Integer): Boolean;
    function FindMergeInfo(ACol, ARow: Integer): TMergeInfo;
    function FindBaseCell(ACol, ARow: Integer): TMergeInfo;

    property MerageInfo[AIndex: Integer]: TMergeInfo read GetMergeInfo;
  end;

  TMiniGrid = class(TStringGrid)
  private
    FColumns: TColumns;
    FMergeInfos: TMergeInfos;
    FActiveColor: TColor;
    FPaintID: Integer;
    FOldCursor: TCursor;
    FUrlCache: TStrings;
    FFixedFont: TFont;

    procedure DoInitColumns;

    procedure Paint; override;
    procedure RepaintCell(ACol, ARow: integer);
    procedure DoDrawAxisLines(ARect: TRect; AState: TGridDrawState);
    procedure DoDrawText(const AText: string; ARect: TRect; AWordBreak: Boolean); overload;
    procedure DoDrawText(const AText: string; ACol, ARow: Integer; ARect: TRect; AWordBreak: Boolean); overload;
    procedure DoDrawUrlText(const AText: string; ARect: TRect);
    procedure WMPaint(var Msg: TWMPAINT); message WM_PAINT;

    function IsUrl(const AText: string): Boolean;
    function PtInUrlRange(const X, Y: Integer): Boolean;
    function CellRect(ACol, ARow: Integer): TRect;
    function GetCellEx(ACol, ARow: Integer): String;
    procedure SetCellEx(ACol, ARow: Integer; const Value: String);
    function GetColCount: Integer;
    procedure SetColCount(const Value: Integer);
    function AnalysisUrl(const ASrcText: string;
      ACol, ARow: Integer): TUrlCache;
    function GetUrlCacheKey(ACol, ARow: Integer): string;
    function FindUrlCache(ACol, ARow: Integer): TUrlCache;

    //隐藏此属性，不让直接访问
    property ColCount: Integer read GetColCount write SetColCount;
  protected
    procedure DrawCell(ACol, ARow: Longint; ARect: TRect;
      AState: TGridDrawState); override;
    function SelectCell(ACol, ARow: Longint): Boolean; override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure MergeCells(ABaseCol, ABaseRow, ARowSpan, AColSpan: Integer);

    property Columns: TColumns read FColumns;
    property MergeInfos: TMergeInfos read FMergeInfos;
    property ActiveColor: TColor read FActiveColor write FActiveColor default clHighlight;
    property PaintId: Integer read FPaintID;
    property Cells[ACol, ARow: Integer]: String read GetCellEx write SetCellEx;
  published
    property FixedFont: TFont read FFixedFont write FFixedFont;
  end;

implementation

/// <summary>
/// 字符串定位函数
/// 因为Delphi6及以下版本并不支持此函数，所以从delphi xe版本迁移代码，省略了asm部分
/// </summary>
function PosEx(const SubStr, S: string; Offset: Integer = 1): Integer;
var
  I, LIterCnt, L, J: Integer;
  PSubStr, PS: PChar;
begin
  if SubStr = '' then
  begin
    Result := 0;
    Exit;
  end;

  { Calculate the number of possible iterations. Not valid if Offset < 1. }
  LIterCnt := Length(S) - Offset - Length(SubStr) + 1;

  { Only continue if the number of iterations is positive or zero (there is space to check) }
  if (Offset > 0) and (LIterCnt >= 0) then
  begin
    L := Length(SubStr);
    PSubStr := PChar(SubStr);
    PS := PChar(S);
    Inc(PS, Offset - 1);

    for I := 0 to LIterCnt do
    begin
      J := 0;
      while (J >= 0) and (J < L) do
      begin
        if (PS + I + J)^ = (PSubStr + J)^ then
          Inc(J)
        else
          J := -1;
      end;
      if J >= L then
      begin
        Result := I + Offset;
        Exit;
      end;
    end;
  end;

  Result := 0;
end;

{ TMiniGrid }

constructor TMiniGrid.Create(AOwner: TComponent);
begin
  inherited;
  FDoubleBuffered := True;
  DefaultDrawing := False;

  FFixedFont := TFont.Create;

  FColumns := TColumns.Create(Self);
  FMergeInfos := TMergeInfos.Create;
  FUrlCache := THashedStringList.Create;

  FixedCols := 1;
  FixedRows := 1;
  RowCount := 2;
  ColWidths[0] := 20;

  FActiveColor := clHighlight;
  FPaintID := 0;
  FOldCursor := Cursor;
end;

destructor TMiniGrid.Destroy;
var
  i: integer;
begin
  if Assigned(FFixedFont) then
    FreeAndNil(FFixedFont);

  if Assigned(FColumns) then
    FreeAndNil(FColumns);

  if Assigned(FMergeInfos) then
    FreeAndNil(FMergeInfos);

  if Assigned(FUrlCache) then
  begin
    for i := FUrlCache.Count - 1 downto 0 do
      FUrlCache.Objects[i].Free;
    FreeAndNil(FUrlCache);
  end;
  inherited;
end;


procedure TMiniGrid.DoInitColumns;
var
  i: Integer;
  objCol: TColumn;
begin
  for i := 0 to FColumns.Count - 1 do
  begin
    objCol := FColumns.Columns[i];
    ColWidths[i+1] := objCol.Width;
    Cols[i+1].Text := objCol.Title;
  end;
end;


function TMiniGrid.CellRect(ACol, ARow: Integer): TRect;
  function _CalcRect(c, r, spanx, spany: Integer): TRect;
  var
    i: Integer;
  begin
    if c < LeftCol then
      c := LeftCol;
    if r < TopRow then
      r := TopRow;

    Result := inherited CellRect(c, r);
    for i := c+1 to (c+spanx) do
      Result.Right := Result.Right + ColWidths[i];

    if c+1 > (c+spanx) then
      Result.Right := Result.Right
    else
      Result.Right := Result.Right + spanx*GridLineWidth;

    for i := r+1 to (r+spany) do
      Result.Bottom := Result.Bottom + RowHeights[i];

    if r+1 > (r+spany) then
      Result.Bottom := Result.Bottom
    else
      Result.Bottom := Result.Bottom + spany*GridLineWidth;
  end;

var
  objMergeInfo: TMergeInfo;
begin
  if not FMergeInfos.IsMergeCell(ACol, ARow) then
    Result := inherited CellRect(ACol, ARow)
  else
  begin
    objMergeInfo := FMergeInfos.FindBaseCell(ACol, ARow);
    if objMergeInfo <> nil then
      Result := _CalcRect(objMergeInfo.Col, objMergeInfo.Row, objMergeInfo.ColSpan, objMergeInfo.RowSpan);
  end;
end;

procedure TMiniGrid.DrawCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  sCellText: string;
  objMI: TMergeInfo;
begin
  Canvas.Font.Assign(Font);

  sCellText := Cells[ACol, ARow];
  if FMergeInfos.IsMergeCell(ACol, ARow) and not FMergeInfos.IsBaseCell(ACol, ARow) then
  begin
    ARect := CellRect(ACol, ARow);
    objMI := FMergeInfos.FindBaseCell(ACol, ARow);
    if (objMI <> nil) and (objMI.PaintId <> FPaintID) then
    begin
      if ((Selection.Left <= (objMI.Col+objMi.ColSpan)) and (Selection.Top <= (objMI.Row+objMI.RowSpan)))
        and ((Selection.Left >= objMI.Col) and (Selection.Top >= objMI.Row)) then
        AState := AState + [gdSelected]
      else
        AState := AState - [gdSelected];

      objMI.PaintId := FPaintID;
      DrawCell(objMI.Col, objMI.Row, ARect, AState);
    end;
    Exit;
  end;

  ARect := CellRect(ACol, ARow);
  if gdFixed in AState then
    Canvas.Brush.Color := FixedColor
  else if gdSelected in AState then
    Canvas.Brush.Color := FActiveColor
  else
    Canvas.Brush.Color := Color;

  Canvas.Pen.Color := Canvas.Brush.Color;
  Canvas.Rectangle(ARect);

  if gdFixed in AState then
  begin
    Canvas.Font.Assign(FFixedFont);
    Canvas.Pen.Color := clGray;
    Canvas.Pen.Width := 1;
  end
  else
  begin
    Canvas.Font.Assign(FColumns.Columns[ACol-1].Font);
    if GridLineWidth > 0 then
      Canvas.Pen.Color := clGray;

    Canvas.Pen.Width := GridLineWidth;
  end;

  DoDrawAxisLines(ARect, AState);
  if TopRow <= ARow then
  begin
    InflateRect(ARect, -GridLineWidth, -GridLineWidth);
    DoDrawText(sCellText, ACol, ARow, ARect, True);
  end;

  if Assigned(OnDrawCell) then
    OnDrawCell(Self, ACol, ARow, ARect, AState);
end;

procedure TMiniGrid.DoDrawAxisLines(ARect: TRect; AState: TGridDrawState);
begin
  if ((goHorzLine in Options) and not (gdFixed in AState)) or
     ((goFixedHorzLine in Options) and (gdFixed in AState)) then
  begin
    Canvas.MoveTo(ARect.Left, ARect.Bottom);
    Canvas.LineTo(ARect.Right, ARect.Bottom);
  end;

  if ((goVertLine in Options) and not (gdFixed in AState)) or
     ((goFixedVertLine in Options) and (gdFixed in AState)) then
  begin
    if UseRightToLeftAlignment then
    begin
      Canvas.MoveTo(ARect.Right,ARect.Bottom);
      Canvas.LineTo(ARect.Right,ARect.Top);
    end
    else
    begin
      Canvas.MoveTo(ARect.Right, ARect.Bottom);
      Canvas.LineTo(ARect.Right, ARect.Top);
    end;
  end;
end;


procedure TMiniGrid.MergeCells(ABaseCol, ABaseRow, ARowSpan, AColSpan: Integer);
var
  mergeInfo: TMergeInfo;
begin
  mergeInfo := TMergeInfo.Create;
  FMergeInfos.Add(mergeInfo);
  mergeInfo.Col := ABaseCol;
  mergeInfo.Row := ABaseRow;
  mergeInfo.RowSpan := ARowSpan;
  mergeInfo.ColSpan := AColSpan;
end;


function TMiniGrid.SelectCell(ACol, ARow: Integer): Boolean;
begin
  RepaintCell(Col,Row);
  Result := inherited SelectCell(Acol, ARow);
  RepaintCell(ACol,ARow);
end;

procedure TMiniGrid.RepaintCell(ACol, ARow: integer);
var
  rc: TRect;
begin
  if HandleAllocated Then 
  begin
    rc := CellRect(ACol, ARow);
    InvalidateRect(Handle, @rc, True);
  end;
end;

procedure TMiniGrid.Paint;
begin
  inherited;
  inc(FPaintID);
end;


/// <summary>
///  绘制Html格式的链接内容
/// </summary>
/// <param name="AText"> 原始文本</param>
/// <param name="ARect"> 内容矩形框</param>
procedure TMiniGrid.DoDrawUrlText(const AText: string; ARect: TRect);
begin
  if AText = '' then
    Exit;

  Canvas.Font.Style := Canvas.Font.Style + [fsUnderline];
  Canvas.Font.Color := clBlue;

  DrawText(Canvas.Handle, PChar(AText), Length(AText), ARect, DT_WORDBREAK or DT_END_ELLIPSIS);

  Canvas.Font.Style := Canvas.Font.Style - [fsUnderline];
  Canvas.Font.Color := Font.Color;
end;


procedure TMiniGrid.DoDrawText(const AText: string; ARect: TRect;
  AWordBreak: Boolean);
var
  rc: TRect;
  oColor: TColor;
  dDrawStyle: DWord;
begin
  rc := ARect;
  oColor := Canvas.Font.Color;
  dDrawStyle := DT_END_ELLIPSIS;
  if AWordBreak then
    dDrawStyle := dDrawStyle or DT_WORDBREAK;
  if IsUrl(AText) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsUnderline];
    Canvas.Font.Color := clBlue;
  end;
  DrawTextEx(Canvas.Handle, PChar(AText), Length(AText), rc, dDrawStyle, nil);

  if IsUrl(AText) then
  begin
    Canvas.Font.Style := Canvas.Font.Style - [fsUnderline];
    Canvas.Font.Color := oColor;
  end;
end;

function TMiniGrid.IsUrl(const AText: string): Boolean;
begin
  Result := (Pos('://', AText) > 0) or (Pos('mailto:', AText) > 0);
end;

/// <summary>
///  解析html格式<A>标签，如下所示
///   exmaple:
///       <A href="http://www.cnblogs.com/5207/">Click here</A>
/// </summary>
/// <param name="ASrcText"></param>
/// <returns></returns>
function TMiniGrid.AnalysisUrl(const ASrcText: string;
  ACol, ARow: Integer): TUrlCache;
var
  oColor: TColor;
  sText, sUrl, sContent: string;
  iPos, iBegin, iEnd: Integer;
begin
  Result := TUrlCache.Create;
  Result.FullData := ASrcText;
  Result.Col := ACol;
  Result.Row := ARow;

  iBegin := 0;
  iBegin := PosEx('<A', ASrcText, 1);
  if iBegin > 0 then
  begin
    iEnd := PosEx('</A>', ASrcText, iBegin);

    sContent := Copy(ASrcText, iBegin, iEnd+Length('</A>'));
    iPos := PosEx('href=', ASrcText, 1);
    if (iPos > 0) then
    begin
      sUrl := '';
      Inc(iPos, Length('href='));
      while ASrcText[iPos] <> '>' do
      begin
        sUrl := sUrl + ASrcText[iPos];
        Inc(iPos);
      end;

      StringReplace(sUrl, '"', '', [rfReplaceAll, rfIgnoreCase]);
      StringReplace(sUrl, '''', '', [rfReplaceAll, rfIgnoreCase]);

      Result.Url := sUrl;
    end;

    iPos := PosEx('>', ASrcText, 1);
    if (iPos > 0) then
    begin
      Inc(iPos);
      sText := '';
      while ASrcText[iPos] <> '<' do
      begin
        sText := sText + ASrcText[iPos];
        Inc(iPos);
      end;

      Result.Text := sText;
    end;
  end
  else if IsUrl(ASrcText) then
  begin
    Result.Url := ASrcText;
    Result.Text := ASrcText;
  end;

  //缓存url资源
  FUrlCache.AddObject(GetUrlCacheKey(ACol, ARow), Result);
end;

function TMiniGrid.PtInUrlRange(const X, Y: Integer): Boolean;
var
  gc: TGridCoord;
  sData: String;
  cellRect: TRect;
  p: TPoint;
  objUrlCache: TUrlCache;
begin
  Result := False;
  gc := Self.MouseCoord(X,Y);
  if (gc.X < 0) or (gc.Y < 0) then
    Exit;
  sData := Self.Cells[gc.X, gc.Y];
  if IsUrl(sData) then
  begin
    objUrlCache := FindUrlCache(gc.X, gc.Y);
    if objUrlCache <> nil then
      sData := objUrlCache.Text;
    
    p.X := X;          
    p.Y := Y;

    cellRect := Self.CellRect(gc.X, gc.Y);
    DrawText(Canvas.Handle, PChar(sData), Length(sData), cellRect, DT_WORDBREAK or DT_CALCRECT);
    Result := (gc.Y > 0) and (gc.Y < Self.RowCount)and PtInRect(cellRect, p);
  end;
end;


procedure TMiniGrid.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  if PtInUrlRange(X, Y) then
    Cursor := crHandPoint
  else
    Cursor := FOldCursor;
end;

procedure TMiniGrid.MouseUp(Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
var
  gc: TGridCoord;
  sUrl: String;
  p: TPoint;
  objUrlCache: TUrlCache;
begin
  inherited;
  if PtInUrlRange(X, Y) then
  begin
    gc := Self.MouseCoord(X, Y);

    objUrlCache := FindUrlCache(gc.X, gc.Y);
    sUrl := Self.Cells[gc.X, gc.Y];
    if objUrlCache <> nil then
      sUrl := objUrlCache.Url;

    ShellExecute(self.handle, 'open', PChar(sUrl), 0, 0, SW_SHOW);
  end;
end;

procedure TMiniGrid.WMPaint(var Msg: TWMPaint);
var
  DC, MemDC: HDC;
  MemBitmap, OldBitmap: HBITMAP;
  PS: TPaintStruct;
begin
  if not FDoubleBuffered or (Msg.DC <> 0) then
  begin
    if not (csCustomPaint in ControlState) and (ControlCount = 0) then
      inherited
    else
      PaintHandler(Msg);
  end
  else
  begin
    DC := GetDC(0);
    MemBitmap := CreateCompatibleBitmap(DC, ClientRect.Right, ClientRect.Bottom);
    ReleaseDC(0, DC);
    MemDC := CreateCompatibleDC(0);
    OldBitmap := SelectObject(MemDC, MemBitmap);
    try
      DC := BeginPaint(Handle, PS);
      Perform(WM_ERASEBKGND, MemDC, MemDC);
      Msg.DC := MemDC;
      WMPaint(Msg);
      Msg.DC := 0;
      BitBlt(DC, 0, 0, ClientRect.Right, ClientRect.Bottom, MemDC, 0, 0, SRCCOPY);
      EndPaint(Handle, PS);
    finally
      SelectObject(MemDC, OldBitmap);
      DeleteDC(MemDC);
      DeleteObject(MemBitmap);
    end;
  end;
end;

function TMiniGrid.GetCellEx(ACol, ARow: Integer): String;
var
  objMI: TMergeInfo;
begin
  if FMergeInfos.IsMergeCell(ACol, ARow) then
  begin
    objMI := FMergeInfos.FindBaseCell(ACol, ARow);
    Result := inherited Cells[objMI.Col, objMi.Row];
  end
  else
    Result := inherited Cells[ACol, ARow];
end;

procedure TMiniGrid.SetCellEx(ACol, ARow: Integer; const Value: String);
var
  objMI: TMergeInfo;
begin
  if FMergeInfos.IsMergeCell(ACol, ARow) then
  begin
    objMI := FMergeInfos.FindBaseCell(ACol, ARow);
    inherited Cells[objMI.Col, objMI.Row] := Value;
    if Assigned(Parent) then
      RepaintCell(ACol, ARow);
  end
  else
  begin
    inherited Cells[ACol, ARow] := Value;
  end;
end;

function TMiniGrid.GetColCount: Integer;
begin
  Result := inherited ColCount;
end;

procedure TMiniGrid.SetColCount(const Value: Integer);
begin
  inherited ColCount := Value+1;

  DoInitColumns;
end;

function TMiniGrid.GetUrlCacheKey(ACol, ARow: Integer): string;
begin
  Result := Format('%d/%d', [ACol, ARow]);
end;

function TMiniGrid.FindUrlCache(ACol, ARow: Integer): TUrlCache;
var
  idx: Integer;
begin
  Result := nil;
  idx := FUrlCache.IndexOf(GetUrlCacheKey(ACol, ARow));
  if idx <> -1 then
    Result := TUrlCache(FUrlCache.Objects[idx]);
end;


procedure TMiniGrid.DoDrawText(const AText: string; ACol, ARow: Integer;
  ARect: TRect; AWordBreak: Boolean);
var
  objUrlCache: TUrlCache;  
begin
  if IsUrl(AText) then
  begin
    objUrlCache := FindUrlCache(ACol, ARow);
    if objUrlCache = nil then
      objUrlCache := AnalysisUrl(AText, ACol, ARow);

    DoDrawUrlText(objUrlCache.Text, ARect);
  end
  else
    DoDrawText(AText, ARect, True);
end;

{ TColumns }

function TColumns.AddColumn: TColumn;
begin
  Result := TColumn.Create(Self);
  Add(Result);

  FOwnerGrid.ColCount := Count;
end;

procedure TColumns.Changed;
begin
  FOwnerGrid.ColCount := Count;
end;

constructor TColumns.Create(AOwner: TMiniGrid);
begin
  inherited Create(True);
  FOwnerGrid := AOwner;
end;

procedure TColumns.DeleteColumn(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex >= Count) then
    Error(@SListIndexError, AIndex);

  self.Delete(AIndex);
  FOwnerGrid.DeleteColumn(AIndex+1);
end;

destructor TColumns.Destroy;
begin

  inherited;
end;

function TColumns.GetColumn(Index: Integer): TColumn;
begin
  if (Index < 0) or (Index >= Count) then
    Error(@SListIndexError, Index);

  Result := TColumn(Items[Index]);
end;

{ TColumn }

constructor TColumn.Create(AOwner: TColumns);
begin
  FOwner := AOwner;
  FWidth := 100;
  FFontColor := clWindowText;
  FHeadBackgroundColor := clWindow;
  FFont := TFont.Create;
end;

destructor TColumn.Destroy;
begin
  inherited;

  if Assigned(FFont) then
    FreeAndNil(FFont);
end;

procedure TColumn.SetFieldName(const Value: string);
begin
  FFieldName := Value;
  FOwner.Changed;
end;

procedure TColumn.SetTitle(const Value: string);
begin
  FTitle := Value;
  FOwner.Changed;
end;

procedure TColumn.SetWidth(const Value: Integer);
begin
  FWidth := Value;
  FOwner.Changed;  
end;

{ TMergeInfos }

function TMergeInfos.FindBaseCell(ACol, ARow: Integer): TMergeInfo;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    with MerageInfo[i] do
    begin
      if (ACol >= Col) and (ACol <= (Col+ColSpan)) and (ARow >= Row) and (ARow <= (Row+RowSpan))then
      begin
        Result := MerageInfo[i];
        Exit;
      end;
    end;
  end;
  Result := nil;
end;

function TMergeInfos.FindMergeInfo(ACol, ARow: Integer): TMergeInfo;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if (MerageInfo[i].Col = ACol) and (MerageInfo[i].Row = ARow) then
    begin
      Result := MerageInfo[i];
      Exit;
    end;
  end;

  Result := nil;
end;

function TMergeInfos.GetMergeInfo(AIndex: Integer): TMergeInfo;
begin
  if (AIndex < 0) or (AIndex >= Count) then
    Error(@SListIndexError, AIndex);

  Result := TMergeInfo(Items[AIndex]);
end;

function TMergeInfos.IsBaseCell(ACol, ARow: Integer): Boolean;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if (MerageInfo[i].Col = ACol) and (MerageInfo[i].Row = ARow) then
    begin
      Result := True;
      Exit;
    end;
  end;

  Result := False;
end;

function TMergeInfos.IsMergeCell(ACol, ARow: Integer): Boolean;
var
  i: Integer;
begin
  Result := False;

  for i := 0 to Count - 1 do
  begin
    with MerageInfo[i] do
    begin
      if (ACol >= Col) and (ACol <= (Col+ColSpan)) and (ARow >= Row) and (ARow <= (Row+RowSpan))then
      begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;


end.
