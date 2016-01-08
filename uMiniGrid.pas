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
  Messages;

type
  TColumn = class(TComponent)
  private
    FTitle: string;
    FFieldName: string;
    FWidth: Integer;
  public
    constructor Create;
    property Title: string read FTitle write FTitle;
    property FieldName: string read FFieldName write FFieldName;
    property Width: Integer read FWidth write FWidth;
  end;

  TColumns = class(TObjectList)
  private
    function GetColumn(Index: Integer): TColumn;
  public
    constructor Create;
    destructor Destroy; override;

    function AddColumn: TColumn;
    procedure DeleteColumn(AIndex: Integer);

    property Columns[Index: Integer]: TColumn read GetColumn;
  end;

  TRow = class(TPersistent)
  public
  end;

  TRows = class(TObjectList)
  public
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
    FRows: TRows;
    FMergeInfos: TMergeInfos;
    FActiveColor: TColor;
    FPaintID: Integer;

    procedure DoInitColumns;
    procedure DoInitRows;

    function CellRect(c, r: Integer): TRect;
    procedure RepaintCell(c, r: integer);
    procedure Paint; override;
    procedure WarpDrawText(const AText: string; ARect: TRect; AWrap: Boolean);
  protected
    procedure DrawColumn(ACol, ARow: Integer; ARect: TRect;
      AState: TGridDrawState);

    procedure DrawCell(ACol, ARow: Longint; ARect: TRect;
      AState: TGridDrawState); override;
    function SelectCell(ACol, ARow: Longint): Boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function AddRow: TRow;

    procedure MergeCells(ACol, ARow, ARowSpan, AColSpan: Integer);

    property Columns: TColumns read FColumns;
    property MergeInfos: TMergeInfos read FMergeInfos;
    property ActiveColor: TColor read FActiveColor write FActiveColor default clBlue;
    property PaintId: Integer read FPaintID;
  published
  end;

implementation

{ TMiniGrid }

constructor TMiniGrid.Create(AOwner: TComponent);
begin
  inherited;

  DefaultDrawing := False;
  FixedCols := 1;
  FixedRows := 1;
  RowCount := 2;
  FColumns := TColumns.Create;
  FRows := TRows.Create;
  FMergeInfos := TMergeInfos.Create;

  FActiveColor := clHighlight;
  FPaintID := 0;
end;

destructor TMiniGrid.Destroy;
begin
  if Assigned(FColumns) then
    FreeAndNil(FColumns);

  if Assigned(FRows) then
    FreeAndNil(FRows);

  if Assigned(FMergeInfos) then
    FreeAndNil(FMergeInfos);
  inherited;
end;

function TMiniGrid.AddRow: TRow;
begin

end;


procedure TMiniGrid.DoInitColumns;
var
  i: Integer;
  objCol: TColumn;
begin
  for i := 0 to FColumns.Count - 1 do
  begin
    objCol := FColumns.Columns[i];
    ColWidths[i] := objCol.Width;
    Cols[i].Text := objCol.Title;
  end;
end;

procedure TMiniGrid.DoInitRows;
begin

end;

function TMiniGrid.CellRect(c, r: Integer): TRect;
var
  i: Integer;
  objMergeInfo: TMergeInfo;
  rc: TRect;
begin
  Result := inherited CellRect(c, r);
  Result.Right := Result.Left + ColWidths[c];
  Result.Bottom := Result.Top + RowHeights[r];
  if FMergeInfos.IsMergeCell(c, r) then
  begin
    objMergeInfo := FMergeInfos.FindBaseCell(c, r);
    if objMergeInfo <> nil then
    begin
      Result := inherited CellRect(objMergeInfo.Col, objMergeInfo.Row);
      for i := 1 to objMergeInfo.ColSpan do
      begin
        rc := inherited CellRect(objMergeInfo.Col+i, r);
        Result.Right := rc.Right;
      end;

      for i := 1 to objMergeInfo.RowSpan do
      begin
        rc := inherited CellRect(c, objMergeInfo.Row+i);
        Result.Bottom := rc.Bottom;
      end;
    end;
  end;
end;

procedure TMiniGrid.DrawColumn(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  sCellText: string;  
begin
  sCellText := Cells[ACol, ARow];
  DrawText(Canvas.Handle, PChar(sCellText), Length(sCellText), ARect, DT_WORDBREAK);
end;

procedure TMiniGrid.DrawCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
var
  sCellText: string;
  objMI: TMergeInfo;
begin
  if Assigned(OnDrawCell) then
    OnDrawCell(Self, ACol, ARow, ARect, AState);

  sCellText := Cells[ACol, ARow];
  if FMergeInfos.IsMergeCell(ACol, ARow) and not FMergeInfos.IsBaseCell(ACol, ARow) then
  begin
    ARect := CellRect(ACol, ARow);
    objMI := FMergeInfos.FindBaseCell(ACol, ARow);
    if (objMI <> nil) and (objMI.PaintId <> FPaintID) then
    begin
      objMI.PaintId := FPaintID;
      DrawCell(objMI.Col, objMI.Row, ARect, AState);
    end;
    Exit;
  end;

  ARect := CellRect(ACol, ARow);
  if gdFixed in AState then
    Canvas.Brush.Color := FixedColor
  else
    Canvas.Brush.Color := Color;
  if gdSelected in AState then
    Canvas.Brush.Color := FActiveColor;

  Canvas.Pen.Color := Canvas.Brush.Color;
  Canvas.Rectangle(ARect);

  if gdFixed in AState then
  begin
    Canvas.Pen.Color := clGray;
    Canvas.Pen.Width := 1;
  end
  else
  begin
    if GridLineWidth > 0 then
      Canvas.Pen.Color := clGray;

    Canvas.Pen.Width := GridLineWidth;
  end;

  WarpDrawText(sCellText, ARect, True);
  
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


procedure TMiniGrid.MergeCells(ACol, ARow, ARowSpan, AColSpan: Integer);
var
  mergeInfo: TMergeInfo;
begin
  mergeInfo := TMergeInfo.Create;
  FMergeInfos.Add(mergeInfo);
  mergeInfo.Col := ACol;
  mergeInfo.Row := ARow;
  mergeInfo.RowSpan := ARowSpan;
  mergeInfo.ColSpan := AColSpan;
end;


function TMiniGrid.SelectCell(ACol, ARow: Integer): Boolean;
begin
  RepaintCell(Col,Row);
  Result := inherited SelectCell(Acol, ARow);
  RepaintCell(ACol,ARow);
end;

procedure TMiniGrid.RepaintCell(c, r: integer);
var
  rc: TRect;
begin
  if HandleAllocated Then 
  begin
    rc := CellRect(c, r);
    InvalidateRect(Handle, @rc, True);
  end;
end;

procedure TMiniGrid.Paint;
begin
  inherited;
  inc(FPaintID);
end;


procedure TMiniGrid.WarpDrawText(const AText: string; ARect: TRect;
  AWrap: Boolean);
var
  rc: TRect;  
begin
  rc := ARect;
  //DrawText(Canvas.Handle, PChar(AText), Length(AText), rc, DT_WORDBREAK or DT_CALCRECT);
  DrawText(Canvas.Handle, PChar(AText), Length(AText), rc, DT_WORDBREAK or DT_LEFT);
end;

{ TColumns }

function TColumns.AddColumn: TColumn;
begin
  Result := TColumn.Create;
  Add(Result);
end;

constructor TColumns.Create;
begin
  inherited Create(True);
end;

procedure TColumns.DeleteColumn(AIndex: Integer);
begin
  if (AIndex < 0) or (AIndex >= Count) then
    Error(@SListIndexError, AIndex);

  self.Delete(AIndex);
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

constructor TColumn.Create;
begin
  FWidth := 100;
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
