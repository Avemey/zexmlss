//****************************************************************
// ZColorStringGrid - компонент,
// представляющий собой StingGrid с возможностью объединения
// ячеек, выравнивание текста по вертикали и горизонтали,
// поворот текста на заданный угол (только TrueType шрифтов) и др.
// Наследован от TStringGrid-а.
// Автор:  Неборак Руслан Владимирович (Ruslan V. Neborak)
// e-mail: avemey(мяу)tut(точка)by
// URL:    http://avemey.com
// Ver:    0.4
// Last update:	 2012.08.04
//----------------------------------------------------------------
// Да, кстати, автор не гарантирует правильную работу компонента и
// не несет ответственности за возможный ущерб в результате
// использования данного программного обеспечения.
//****************************************************************

unit ZColorStringGrid;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, stdctrls, zcftext;

type
  TVerticalAlignment = (vaTop, vaCenter, vaBottom);

  //Для сохранения pen и brush
  TPenBrush = record
    PColor: Tcolor;
    Width : Integer;
    Mode  : TPenMode;
    PStyle: TPenStyle;
    BColor: TColor;
    BStyle: TBrushStyle;
  end;

  TBorderCellStyle = (sgLowered, sgRaised, sgNone);
  // sgLowered  - "эдитор"
  // sgRaised   - "кнопка"
  // sgNone     - без рамки

  TDrawMergeCellEvent = procedure (Sender: TObject; ACol, ARow: Longint;
    Rect: TRect; State: TGridDrawState; var CellCanvas: TCanvas) of object;

  TZColorStringGrid = Class;

  //выделенная ячейка
  TSelectColor = class(TPersistent)
  private
    FGrid: TZColorStringGrid;
    FBGColor : TColor;                   //фон выделеной ячейки
    FFontColor: TColor;                  //цвет шрифта в выделенной ячейке
    FColoredSelect: boolean;             //использовать ли цветную раскраску?
    FUseFocusRect: boolean;              //использовать ли рамку на ячейке с фокусом ввода?
    procedure SetBGColor(const Value: TColor);
    procedure SetColoredSelect(const Value: Boolean);
    procedure SetUseFocusRect(const Value: Boolean);
  public
    constructor Create(AGrid: TZColorStringGrid); virtual;
  published
    property BGColor : TColor read FBGColor write SetBGColor default clHighLight;
    property FontColor: TColor read FFontColor write FFOntColor default clWhite;
    property ColoredSelect: boolean read FColoredSelect write SetColoredSelect default true;
    property UseFocusRect: boolean read FUseFocusRect write SetUseFocusRect default true;
  end;

  //линии
  TLineDesign = class(TPersistent)
  private
    FGrid: TZColorStringGrid;
    FLineColor: TColor;
    FLineUpColor: TColor;
    FLineDownColor: TColor;
    procedure SetLineColor(const Value: TColor);
    procedure SetLineUpColor(const Value: TColor);
    procedure SetLineDownColor(const Value: TColor);
  public
    constructor Create(AGrid: TZColorStringGrid); virtual;
  published
    property LineColor: TColor read FLineColor write SetLineColor default clGray;
    property LineUpColor: TColor read FLineUpColor write SetLineUpColor default clGray;
    property LineDownColor: TColor read FLineDownColor write SetLineDownColor default clBlack;
  end;

  //"стиль" ячейки
  TCellStyle = Class(TPersistent)
  private
    FGrid: TZColorStringGrid;
    FFont: TFont;
    FBGColor: TColor;                           //Цвет фона
    FHorizontalAlignment: TAlignment;           //горизонтаьное выравнивание
    FVerticalAlignment: TVerticalAlignment;     //вертикальное выравнивание
    FWordWrap: boolean;
    FSizingHeight: boolean;
    FSizingWidth: boolean;
    FBorderCellStyle: TBorderCellStyle;         //стиль рамки ячейки
    FIndentH: byte;                             //Отступ по горизонтали
    FIndentV: byte;                             //Отступ по вертикали
    FRotate: integer;                           //Угол поворота текста
    procedure SetFont(const Value: TFont);
    procedure SetBGColor(const Value: TColor);
    procedure SetHAlignment(const Value: TAlignment);
    procedure SetVAlignment(const Value: TVerticalAlignment);
    procedure SetBorderCellStyle(const Value: TBorderCellStyle);
    procedure SetIndentV(const Value: byte);
    procedure SetindentH(const Value: byte);
    procedure SetRotate(const Value: integer);
  public
    constructor Create(AGrid: TZColorStringGrid); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    procedure AssignTo(Source: TPersistent); override;
  published
    property Font: Tfont read FFont write SetFont;
    property BGColor: TColor read FBGColor write SetBGColor;
    property BorderCellStyle: TBorderCellStyle read FBorderCellStyle write SetBorderCellStyle default sgNone;
    property IndentH: byte read FIndentH write SetIndentH default 2;
    property IndentV: byte read FIndentV write SetIndentV default 0;
    property VerticalAlignment: TVerticalAlignment read FVerticalAlignment write SetVAlignment default vaCenter;
    property HorizontalAlignment: TAlignment read FHorizontalAlignment write SetHAlignment default taLeftJustify;
    property Rotate: integer read FRotate write SetRotate default 0;
    property SizingHeight: boolean read FSizingHeight write FSizingHeight default false;
    property SizingWidth: boolean read FSizingWidth write FSizingWidth default false;
    property WordWrap: boolean read FWordWrap write FWordWrap default False;
  end;

  TCellLine = array of TCellStyle;

  //Объединённые ячейки
  TMergeCells = Class
  private
    FGrid: TZColorStringGrid;
    FCount: integer;
    FMergeArea: Array of TRect;
    function Get_Item(Num: Integer): TRect;
  protected
  public
    constructor Create(AGrid: TZColorStringGrid); virtual;
    destructor Destroy(); override;
    procedure Clear();
    function AddRect(Rct:TRect): byte;
    function AddRectXY(x1,y1,x2,y2: integer): byte;
    function DeleteItem(num: integer): boolean;
    function InLeftTopCorner(ACol, ARow: integer): integer;
    function InMergeRange(ACol, ARow: integer): integer;
    function GetWidthArea(num: integer): integer;
    function GetHeightArea(num: integer): integer;
    function GetSelectedArea(SetSelected: boolean): TGridRect;
    function GetMergeYY(ARow: integer): TPoint;
    property Count: integer read FCount;
    property Items[Num: Integer]: TRect read Get_Item;
  end;

  //вместо стандартного InplaceEditor-а
  TZInplaceEditor = class(TMemo)
  private
    FGrid: TZColorStringGrid;
    FExEn: integer;  // "костыль" для исправной работы DoEnter/DoExit *^_^*
  protected
    procedure DoEnter; override;
    procedure DoExit; override;
    procedure Change; override;
    procedure DblClick; override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyPress(var Key: Char); override;
    procedure KeyUp(var Key: Word; Shift: TShiftState); override;
  public
    constructor Create(AOwner: TComponent); override;
    function  DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint): Boolean; override;
    function  DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint): Boolean;override;
  end;

  TInplaceEditorOptions = class (TPersistent)
  private
    FGrid: TZColorStringGrid;
    FFontColor: TColor;
    FBGColor: TColor;
    FBorderStyle: TBorderStyle;
    FAlignment: TAlignment;
    FWordWrap: Boolean;
    FUseCellStyle: Boolean;
    procedure SetFontColor(const Value: TColor);
    procedure SetBGColor(const Value: TColor);
    procedure SetBorderStyle(const Value: TBorderStyle);
    procedure SetAlignment(const Value: TAlignment);
    procedure SetWordWrap(const Value: Boolean);
    procedure SetUseCellStyle(const Value: Boolean);
  protected
  public
    constructor Create(AGrid: TZColorStringGrid); virtual;
  published
    property FontColor: TColor read FFontColor write SetFontColor default clblack;
    property BGColor: TColor read FBGColor write SetBGColor default clWhite;
    property BorderStyle: TBorderStyle read FBorderStyle write SetBorderStyle default bsNone;
    property Alignment: TAlignment read FAlignment write SetAlignment default taLeftJustify;
    property WordWrap: Boolean read FWordWrap write SetWordWrap default true;
    property UseCellStyle: Boolean read FUseCellStyle write SetUseCellStyle default true;
  end;

  TZColorStringGrid = class(TStringGrid)
  private
    FOnBeforeTextDrawCell: TDrawCellEvent;
    FOnBeforeTextDrawCellMerge: TDrawMergeCellEvent;
    FOnDrawCellMerge: TDrawMergeCellEvent;
    FCellStyle: array of TCellLine;     //стили ячеек
    FMergeCells: TMergeCells;           //объединённые ячейки;
    FDefaultFixedCellParam: TCellStyle; //стиль фиксированных ячеек по-умолчанию
    FDefaultCellParam: TCellStyle;      //стиль ячеек по-умолчанию
    FLineDesign : TLineDesign;          //линии
    FWordWrap: boolean;
    FUseCellWordWrap: boolean;
    //для того, чтобы старый эдитор не выскакивал, было переопределено свойство Options
    //и добавлено поле FgoEditing
    FOptions: TGridOptions;             //настройки
    FgoEditing: boolean;                //для goEditing
    //увеличивать длину/высоту столбца/строки
    FSizingHeight: boolean;
    FSizingWidth: boolean;
    FUseCellSizingHeight: boolean;
    FUseCellSizingWidth: boolean;
    FSelectedColors: TSelectColor;
    FUseBitMap: boolean;
    FRctBitmap: TRect;
    FBmp: TBitMap;                      //для рисования объединённых ячеек
    FRctRectangle: TRect;
    FPenBrush: TPenBrush;               //хроним настройки pen и brush

    FZInplaceEditor: TZInplaceEditor;
    FInplaceEditorOptions: TInplaceEditorOptions;

    FDontDraw: boolean;

    FMouseXY : TGridCoord;

    FLastCell: array [0..1,0..6] of Integer;

    procedure SetCellStyle(ACol, ARow: integer; const value: TCellStyle);
    function  GetCellStyle(ACol, ARow: integer): TCellStyle;
    procedure SetCellStyleRow(ARow: integer; fixedCol: boolean; const value: TCellStyle);
    procedure SetCellStyleCol(ACol: integer; fixedRow: boolean; const value: TCellStyle);
    //ячейки/фиксированные ячейки по умолчанию
    procedure SetDefaultCellParam(const Value: TCellStyle);
    function  GetDefaultCellParam: TCellStyle;
    procedure SetDefaultFixedCellParam(const Value: TCellStyle);
    function  GetDefaultFixedCellParam: TCellStyle;
    //RowCount / ColCount
    procedure SetRowCount(const Value: integer);
    function  GetRowCount: integer;
    procedure SetColCount(const Value: integer);
    function  GetColCount: integer;
    //FixedColor
    procedure SetFixedColor(const Value: TColor);
    function  GetFixedColor: TColor;
    //FixedCols / FixedRows
    procedure SetFixedCols(const value: integer);
    function  GetFixedCols(): integer;
    procedure SetFixedRows(const value: integer);
    function  GetFixedRows(): integer;

    procedure SetCells(ACol, ARow: Integer; const Value: string);
    function  GetCells(ACol, ARow: Integer): string;
    procedure Initialize();
    procedure StyleRowMove (FromIndex, ToIndex: integer);
    procedure GetCalcRect(var ARect: TRect; var frmparams: integer; ACol, ARow: integer; var NumMergeArea: integer);
    procedure SetCellCount(oldColCount, newColCount, oldRowCount, newRowCount: integer);

    procedure SetEditorMode(Value: boolean);
    function  GetEditorMode: Boolean;
    procedure SavePenBrush(var CellCanvas: TCanvas);
    Procedure RestorePenBrush(var CellCanvas: TCanvas);
    function  GetOptions: TGridOptions;
    procedure SetOptions(Value: TGridOptions);
    procedure GetParamsForText(ACol, ARow: integer; var frmparams: integer; var AlignmentHorizontal, AlignmentVertical: integer; var ZWordWrap: boolean);
  protected
    procedure CalcCellSize(ACol, ARow: integer); virtual;
    procedure ColumnMoved(FromIndex, ToIndex: Longint); override;
    procedure DrawCell(ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState); override;
    procedure DblClick; override;
    procedure DeleteColumn(ACol: integer); override;
    procedure DeleteRow(ARow: integer); override;
    procedure DoEnter; override;
    procedure DoExit; override;
    function  DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint): Boolean; override;
    function  DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint): Boolean;override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyPress(var Key: Char); override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure Paint; override;
    procedure RowMoved(FromIndex, ToIndex: Longint); override;
    procedure RowSelectYY(key: word); virtual;
    procedure SetEditText(ACol, ARow: Longint; const Value: string);override;
    function  SelectCell(ACol, ARow: Longint): Boolean; override;
    procedure ShowEditor; virtual;
    procedure ShowEditorFirst; virtual;
    procedure TopLeftChanged; override;
    procedure HideEditor; virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy();override;
    property DontDraw: boolean read FDontDraw write FDontDraw;
    property Cells[ACol, ARow: Integer]: string read GetCells write SetCells;
    property CellStyle[ACol, ARow: integer]: TCellStyle read GetCellStyle write SetCellStyle;
    property CellStyleRow[ARow: integer; fixedCol: boolean]: TCellStyle write SetCellStyleRow; //только на запись!
    property CellStyleCol[ACol: integer; fixedRow: boolean]: TCellStyle write SetCellStyleCol; //только на запись!
    property EditorMode: Boolean read GetEditorMode write SetEditorMode;
    property MergeCells: TMergeCells read FMergeCells write FMergeCells;
    property ZInplaceEditor: TZInplaceEditor read FZInplaceEditor write FZInplaceEditor;
  published
    property ColCount: integer read GetColCount write SetColCount default 5;
    property DefaultCellStyle: TCellStyle read GetDefaultCellParam write SetDefaultCellParam;
    property DefaultFixedCellStyle: TCellStyle read GetDefaultFixedCellParam write SetDefaultFixedCellParam;
    property FixedColor: TColor read GetFixedColor write SetFixedColor;// default clBtnFace;
    property FixedCols: Integer read GetFixedCols write SetFixedCols default 1;
    property FixedRows: Integer read GetFixedRows write SetFixedRows default 1;
    property LineDesign: TLineDesign read FLineDesign write FLineDesign;
    property OnBeforeTextDrawCell: TDrawCellEvent read FOnBeforeTextDrawCell write FOnBeforeTextDrawCell;
    property OnBeforeTextDrawMergeCell: TDrawMergeCellEvent read FOnBeforeTextDrawCellMerge write FOnBeforeTextDrawCellMerge;
    property OnDrawMergeCell: TDrawMergeCellEvent read FOnDrawCellMerge write FOnDrawCellMerge;
    property Options: TGridOptions read GetOptions write SetOptions
      default [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine,
      goRangeSelect];
    property RowCount: integer read GetRowCount write SetRowCount default 5;
    property SelectedColors: TSelectColor read FSelectedColors write FSelectedColors;
    property SizingHeight: boolean read FSizingHeight write FSizingHeight default false;
    property SizingWidth: boolean read FSizingWidth write FSizingWidth default false;
    property UseCellSizingHeight: boolean read FUseCellSizingHeight write FUseCellSizingHeight default false;
    property UseCellSizingWidth: boolean read FUseCellSizingWidth write FUseCellSizingWidth default false;
    property UseCellWordWrap: boolean read FUseCellWordWrap write FUseCellWordWrap default false;
    property WordWrap: boolean read FWordWrap write FWordWrap default false;
    property InplaceEditorOptions: TInplaceEditorOptions read FInplaceEditorOptions write FInplaceEditorOptions;
  end;

procedure Register;

implementation

//::: TInplaceEditorOptions :::
constructor TInplaceEditorOptions.Create(AGrid: TZColorStringGrid);
begin
  inherited Create;
  FGrid := AGrid;
  FontColor := ClBlack;
  BGColor := ClWhite;
  Borderstyle := bsNone;
  Alignment := taLeftJustify;
  WordWrap := true;
  UseCellStyle := true;
end;

//установить цвет шрифта в ZInplaceEditor-е
procedure TInplaceEditorOptions.SetFontColor(const Value: TColor);
begin
  if Value <> FFontColor then
  begin
    FFontColor := Value;
    if not FUseCellStyle then
    //на случай, если ZInplaceEditor не существует
    if assigned(FGrid.ZInplaceEditor) then
    FGrid.ZInplaceEditor.Font.Color := FFontColor;
  end;
end;

//установить цвет фона ZInplaceEditor-а
procedure TInplaceEditorOptions.SetBGColor(const Value: TColor);
begin
  if Value <> FBGColor then
  begin
    FBGColor := Value;
    if not FUseCellStyle then
    //на случай, если ZInplaceEditor не существует
    if assigned(FGrid.ZInplaceEditor) then
    FGrid.ZInplaceEditor.Color := FBGColor;
  end;
end;

//установить BorderStyle у ZInplaceEditor-а
procedure TInplaceEditorOptions.SetBorderStyle(const Value: TBorderStyle);
begin
  if Value <> FBorderStyle then
  begin
    FBorderStyle := Value;
    if not FUseCellStyle then
    //на случай, если ZInplaceEditor не существует
    if assigned(FGrid.ZInplaceEditor) then
    FGrid.ZInplaceEditor.BorderStyle := FBorderStyle;
  end;
end;

//Установить выравнивание
procedure TInplaceEditorOptions.SetAlignment(const Value: TAlignment);
begin
  if Value <> FAlignment then
  begin
    FAlignment := Value;
    if not FUseCellStyle then
    if assigned(FGrid.ZInplaceEditor) then
    FGrid.ZInplaceEditor.Alignment := FAlignment;
  end;
end;

procedure TInplaceEditorOptions.SetWordWrap(const Value: boolean);
begin
  if Value <> FWordWrap then
  begin
    FwordWrap := Value;
    if not FUseCellStyle then
    if assigned(FGrid.ZInplaceEditor) then
    FGrid.ZInplaceEditor.WordWrap := FWordWrap;
  end;
end;

procedure TInplaceEditorOptions.SetUseCellStyle(const Value: boolean);
begin
  if Value <> FUseCellStyle then
  begin
    FUseCellStyle := Value;
    if not FUseCellStyle then
    if assigned(FGrid.ZInplaceEditor) then
    begin
      FGrid.ZInplaceEditor.WordWrap := FWordWrap;
      FGrid.ZInplaceEditor.BorderStyle := FBorderStyle;
      FGrid.ZInplaceEditor.Alignment := FAlignment;
      FGrid.ZInplaceEditor.Color := FBGColor;
      FGrid.ZInplaceEditor.Font.Color := FFontColor;
    end;
  end;
end;

//::: TZInplaceEditor :::
constructor TZInplaceEditor.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  DoubleBuffered := true;
  TabStop := false;
  FGrid := AOwner as TZColorStringGrid;
  parent := FGrid;
  BorderStyle := bsNone;
  visible := false;
end;

procedure TZInplaceEditor.Change;
begin
  FGrid.Cells[FGrid.col, FGrid.row] := lines.Text;
end;

procedure TZInplaceEditor.DblClick();
begin
  FGrid.Dblclick;
end;

procedure TZInplaceEditor.DoEnter();
begin
  FExEn := 1;
end;

procedure TZInplaceEditor.DoExit();
begin
  if (FGrid.FLastCell[1,0]=FGrid.Col)and(FGRid.FLastCell[1,1]=FGRID.Row) then
    FGRID.FLastCell[1,0] := -77;
  if (not self.FGrid.Focused)and(FExEn <> 33) then
  begin
    FExEn := 3;
    self.FGrid.DoExit;
    self.FGrid.HideEditor;
    FExEn := 0;
  end;
end;

function  TZInplaceEditor.DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint): Boolean;
begin
  FExEn := 33;
  FGrid.HideEditor;
  Result := FGrid.DoMouseWheelDown(Shift, MousePos);
  if (goEditing in FGrid.Options)and(not (goRowSelect in FGrid.Options))and (goAlwaysShowEditor in FGrid.Options) then
    FGrid.ShowEditorFirst
  else
    if FGrid.CanFocus then FGrid.SetFocus;
end;

function  TZInplaceEditor.DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint): Boolean;
begin
  FExEn := 33;
  FGrid.HideEditor;
  Result := FGrid.DoMouseWheelUp(Shift, MousePos);
  if (goEditing in FGrid.Options)and(not (goRowSelect in FGrid.Options))and (goAlwaysShowEditor in FGrid.Options) then
    FGrid.ShowEditorFirst
  else
    if FGrid.CanFocus then FGrid.SetFocus;
end;

procedure TZInplaceEditor.KeyDown(var Key: Word; Shift: TShiftState);

  procedure SendToGrid;
  begin
    FExEn := 33;
    FGrid.KeyDown(Key, Shift);
    if (goAlwaysShowEditor in FGrid.Options)and (goEditing in FGrid.Options) and (not(goRowSelect in FGrid.Options)) then
      FGrid.ShowEditorFirst
    else
      if FGrid.CanFocus then FGrid.SetFocus;
    key := 0;
  end;

begin
  case Key of
    VK_PRIOR, VK_NEXT: SendToGrid;  
    VK_DELETE: if not FGrid.CanEditModify then key := 0; //delete
    VK_UP: if CaretPos.y < 1 then SendToGrid;
    VK_DOWN: if CaretPos.y >= Lines.Count - 1 then SendToGrid;
    VK_LEFT: if selStart <1 then SendToGrid;
    VK_RIGHT: if SelStart >= length(Lines.Text) then SendToGrid;
  end;
  if Key <>0 then
  begin
    inherited KeyDown(Key, Shift);
  end;
end;

procedure TZInplaceEditor.KeyPress(var Key: Char);
begin
  if key=#13 then
  begin
    FGrid.Cells[FGrid.col, FGrid.row] := lines.Text;
    FGrid.KeyPress(key);
    Key := #0;
    if (goEditing in FGrid.Options)and(not (goRowSelect in FGrid.Options))and (goAlwaysShowEditor in FGrid.Options) then else
    begin
      FGrid.HideEditor;
      FGrid.SetFocus;
    end;
    exit;
  end;
  FGrid.KeyPress(key);
  if not FGrid.CanEditAcceptKey(key) then key := #0;
  case key of
    #9, #27: key := #0;
    else
      if not FGrid.CanEditModify then key := #0;
  end;
  if key <> #0 then
    inherited KeyPress(key);
end;

procedure TZInplaceEditor.KeyUp(var Key: Word; Shift: TShiftState);
begin
  FGrid.KeyUp(Key, Shift);
  inherited KeyUp(key, Shift);
end;

//::: TSelectColor :::
constructor TSelectColor.Create(AGrid: TZColorStringGrid);
begin
  inherited Create;
  FGrid := AGrid;
  FBGColor := clHighLight;
  FFontColor := clWhite;
  FColoredSelect := true;
  FUseFocusRect := true;
end;

procedure TSelectColor.SetBGColor(const Value: TColor);
begin
  if FBGColor <> value then
  begin
    FBGColor := value;
    FGrid.Invalidate;
  end;  
end;

procedure TSelectColor.SetColoredSelect(const Value: Boolean);
begin
  if FColoredSelect <> value then
  begin
    FColoredSelect := value;
    FGrid.Invalidate;
  end;
end;

procedure TSelectColor.SetUseFocusRect(const Value: Boolean);
begin
  if value <> FUseFocusRect then
  begin
    FUseFocusRect := value;
    FGrid.Invalidate;
  end;
end;

//::: TLineDesign :::
constructor TLineDesign.Create(AGrid: TZColorStringGrid);
begin
  inherited Create;
  FGrid := AGrid;
  FLineColor := clGray;
  FLineUpcolor := clWhite;
  FLineDownColor := clBlack;
end;

procedure TLineDesign.SetLineColor(const Value: TColor);
begin
  if FLineColor <> Value then
  begin
    FLineColor := Value;
    FGrid.Invalidate;
  end;
end;

procedure TLineDesign.SetLineUpColor(const Value: TColor);
begin
  if FLineUpColor <> Value then
  begin
    FLineUpColor := Value;
    FGrid.Invalidate;
  end;
end;

procedure TLineDesign.SetLineDownColor(const Value: TColor);
begin
  if FLineDownColor <> Value then
  begin
    FLineDownColor := Value;
    FGrid.Invalidate;
  end;
end;

//::: TCellStyle :::
constructor TCellStyle.Create(AGrid: TZColorStringGrid);
begin
  inherited Create;
  FFont := TFont.Create;
  if assigned(AGrid) then
  begin
    FFont.Assign(AGrid.font);
    FGrid := AGrid;
  end;
  FBorderCellStyle := sgNone;
  self.VerticalAlignment := vaCenter;
  self.HorizontalAlignment := taLeftJustify;
  self.FIndentH := 2;
  self.FIndentV := 0;
  Self.FRotate := 0;
end;

destructor TCellStyle.Destroy();
begin
  FFont.Free;
  inherited;
end;

procedure TCellStyle.SetFont(const Value: TFont);
begin
  FFont.Assign(Value);
  if assigned(FGrid) then
  begin
    FGrid.fLastCell[1,0] := -29;
    FGrid.CalcCellSize(FGrid.Col, FGrid.Row);
    FGrid.Invalidate;
  end;
end;

procedure TCellStyle.SetBGColor(Const Value: TColor);
begin
  FBGColor := Value;
  if assigned(FGrid) then
    FGrid.Invalidate;
end;

procedure TCellStyle.AssignTo(Source: TPersistent);
var
  zSource: TCellStyle;

begin
  if (Source is TCellStyle) then
  begin
    zSource := (Source as TCellStyle);
    zSource.FBGColor := FBGColor;
    zSource.FRotate := FRotate;
    zSource.FFont.Assign(FFont);
    zSource.FHorizontalAlignment := FHorizontalAlignment;
    zSource.FVerticalAlignment := FVerticalAlignment;
    zSource.FWordWrap := FWordWrap;
    zSource.FSizingHeight := FSizingHeight;
    zSource.FSizingWidth := FSizingWidth;
    zSource.FBorderCellStyle := FBorderCellStyle;
    zSource.FIndentH := FIndentH;
    zSource.FIndentV := FIndentV;
  end else
  if (Source is TFont) then
    (Source as TFont).assign(FFont)
  else
    inherited;
end;

procedure TCellStyle.Assign(Source: TPersistent);
var
  zSource: TCellStyle;

begin
  if (Source is TCellStyle) then
  begin
    zSource := (Source as TCellStyle);
    FBGColor := zSource.FBGColor;
    FRotate := zSource.FRotate;
    FFont.Assign(zSource.FFont);
    FHorizontalAlignment := zSource.FHorizontalAlignment;
    FVerticalAlignment := zSource.FVerticalAlignment;
    FWordWrap := zSource.FWordWrap;
    FSizingHeight := zSource.FSizingHeight;
    FSizingWidth := zSource.FSizingWidth;
    FBorderCellStyle := zSource.FBorderCellStyle;
    FIndentH := zSource.FIndentH;
    FIndentV := zSource.FIndentV;
    if assigned(FGrid) then
      FGrid.fLastCell[1,0] := -25;
  end else
  if (Source is TFont) then
    FFont.Assign(Source as TFont)
  else
    inherited;
end;

procedure TCellStyle.SetBorderCellStyle(const Value: TBorderCellStyle);
begin
  if FBorderCellStyle <> value then
  begin
    FBorderCellStyle := value;
    if assigned(FGrid) then
    FGrid.invalidate;
  end;
end;

procedure TCellstyle.SetHAlignment(const Value: TAlignment);
begin
  if FHorizontalAlignment <> Value then
  begin
    FHorizontalAlignment := Value;
    if assigned(FGrid) then
    begin
      FGrid.fLastCell[1,0] := -22;
      FGrid.Invalidate;
    end;
  end;
end;

procedure TCellStyle.SetVAlignment(const Value: TVerticalAlignment);
begin
  if FVerticalAlignment <> Value then
  begin
    FVerticalAlignment := Value;
    if assigned(FGrid) then
    begin
      FGrid.fLastCell[1,1] := -23;
      FGrid.Invalidate;
    end;
  end;
end;

procedure TCellStyle.SetIndentV(const Value: byte);
begin
  if (FIndentV <> Value) then
  begin
    FIndentV := value;
    if (Assigned(FGrid)) then
    begin
      FGrid.fLastCell[1, 1] := -23;
      FGrid.Invalidate;
    end;
  end;
end;

procedure TCellStyle.SetindentH(const Value: byte);
begin
  if (FIndentH <> Value) then
  begin
    FIndentH := value;
    if (Assigned(FGrid)) then
    begin
      FGrid.fLastCell[1, 1] := -23;
      FGrid.Invalidate;
    end;
  end;
end;

procedure TCellStyle.SetRotate(const Value: integer);
begin
  if (FRotate <> Value) then
  begin
    FRotate := Value;
    if (Assigned(FGrid)) then
    begin
      FGrid.fLastCell[1, 1] := -23;
      FGrid.Invalidate;
    end;
  end;
end;

//::: TMergeCells :::

constructor TMergeCells.Create(AGrid: TZColorStringGrid);
begin
  inherited Create;
  FGrid := AGrid;
  FCount := 0;
end;

destructor TMergeCells.destroy();
begin
  Clear();
  inherited;
end;

procedure TMergeCells.Clear();
begin
  setlength(FMergeArea, 0);
  FCount := 0;
end;

function TMergeCells.GetSelectedArea(SetSelected: boolean): TGridRect;
var
  a: array of integer;
  selarea: TGridRect;

  function LineInSelect(x1, y1, x2, y2: integer): boolean;
  begin
    {t}
    if (((x1 >= selarea.Left) and (x1 <= selarea.Right)) or ((x2 >= selarea.Left) and (x2 <= selarea.Right))) and
       ((y1 <= selarea.top) and (y2 >= selarea.top))
      then result := true
    else
    {b}
    if (((x1 >= selarea.Left) and (x1 <= selarea.Right)) or ((x2 >= selarea.Left) and (x2 <= selarea.Right))) and
       ((y1 <= selarea.bottom) and (y2 >= selarea.bottom))
      then result := true
    else
    {l}
    if (((y1 >= selarea.top) and (y1 <= selarea.bottom)) or ((y2 >= selarea.top) and (y2 <= selarea.bottom))) and
       ((x1 <= selarea.left) and (x2 >= selarea.left))
      then result := true
    else
    {r}
    if (((y1 >= selarea.top) and (y1 <= selarea.bottom)) or ((y2 >= selarea.top) and (y2 <= selarea.bottom))) and
       ((x1 <= selarea.right) and (x2 >= selarea.right))
      then result := true
    else
      result := false;
  end;

  procedure FindArea;
  var
    i: integer;

  begin
    for i := 0 to count - 1 do
    begin
      if a[i] = 0 then
      begin
        if LineInSelect(Get_Item(i).Left, Get_Item(i).Top, Get_Item(i).Left, Get_Item(i).bottom) or
           LineInSelect(Get_Item(i).Left, Get_Item(i).Top, Get_Item(i).Right, Get_Item(i).top) or
           LineInSelect(Get_Item(i).right, Get_Item(i).Top, Get_Item(i).right, Get_Item(i).bottom) or
           LineInSelect(Get_Item(i).Left, Get_Item(i).bottom, Get_Item(i).right, Get_Item(i).bottom) then
        begin
          a[i] := 1;
          if selarea.Left > Get_Item(i).Left then selarea.Left := Get_Item(i).Left;
          if selarea.Right < Get_Item(i).Right then selarea.Right := Get_Item(i).right;
          if selarea.Top > Get_Item(i).Top then selarea.Top := Get_Item(i).Top;
          if selarea.Bottom < Get_Item(i).Bottom then selarea.Bottom := Get_Item(i).Bottom;
          FindArea;
        end;
      end;
    end;
  end;

begin
  selarea := FGrid.Selection;
  if goRowSelect in FGrid.Options then
  begin
    selarea.Left := FGrid.FixedCols;
    selarea.Right := FGrid.ColCount - 1;
  end;
  if Count = 0 then
  begin
    result := selarea;
    exit;
  end;
  setlength(a, count);
  FindArea;
  setlength(a, 0);
  if SetSelected then FGrid.Selection := selarea;
  result := selarea;
end;

function TMergeCells.GetMergeYY(ARow: integer): TPoint;
var
  a: array of byte;
  MinYMaxY: TPoint;

  procedure FindAllMerges(y1, y2: integer);
  var
    i: integer;

  begin
    if MinYMaxY.x > y1 then MinYMaxY.x := y1;
    if MinYMaxY.y < y2 then MinYMaxY.y := y2;
    for i := 0 to count - 1 do
    begin
      if a[i] = 0 then
      begin
        if ((y2 >= FMergeArea[i].top)and(y2 <= FMergeArea[i].Bottom ))or
           ((y1 >= FMergeArea[i].top)and(y1 <= FMergeArea[i].Bottom )) then
        begin
          a[i] := 1;
          FindAllMerges(FMergeArea[i].Top, FMergeArea[i].Bottom);
        end;
      end;
    end;
  end;

begin
  MinYMaxY.x := ARow;
  MinYMaxY.y := ARow;
  if Count > 0 then
  begin
    setlength(a, count);
    FindAllMerges(ARow, ARow);
    setlength(a, 0);
  end;
  result := MinYMaxY;
end;

//добавить "объединённую" область
//принимает прямоугольник (Rct: TRect)
//возвращает:
//      0 - всё нормально, область добавилась
//      1 - указанная область выходит за границы грида, область не добавлена
//      2 - указанная область пересекается(входит) в введённые области , область не добавлена
//      3 - область из одной ячейки не добавляет
//      4 - попытка объединить фиксированные и нефиксированные ячейки
function TMergeCells.AddRect(Rct: TRect): byte;
var
  i: integer;

function usl(rct1, rct2: TRect): boolean;
begin
  result := (((rct1.Left >= rct2.Left) and (rct1.Left <= rct2.Right)) or
      ((rct1.right >= rct2.Left) and (rct1.right <= rct2.Right))) and
     (((rct1.Top >= rct2.Top) and (rct1.Top <= rct2.Bottom)) or
      ((rct1.Bottom >= rct2.Top) and (rct1.Bottom <= rct2.Bottom)));
end;

begin
  if (rct.Left<0)or(rct.Top<0)or(rct.Right>FGrid.ColCount-1)or
     (rct.Bottom>FGrid.RowCount-1) or (rct.left>FGrid.ColCount-1) or
     (rct.top>FGrid.RowCount-1) then
  begin
    result := 1;
    exit;
  end;
  if (rct.Right - rct.Left = 0)and(rct.Bottom - rct.Top = 0) then
  begin
    result := 3;
    exit;
  end;
  for i := 0 to Count-1 do
  if {(((FMergeArea[i].Left>=rct.Left) and (FMergeArea[i].Left<=rct.Right))or((FMergeArea[i].right>=rct.Left) and (FMergeArea[i].right<=rct.Right)))and
     (((FMergeArea[i].Top>=rct.Top) and (FMergeArea[i].Top<=rct.Bottom))or((FMergeArea[i].Bottom>=rct.Top) and (FMergeArea[i].Bottom<=rct.Bottom)))}
     usl(FMergeArea[i], rct) or usl(rct, FMergeArea[i]) then
  begin
    result := 2;
    exit;
  end;
  if ((rct.Left<FGrid.FixedCols) and (rct.Right>=FGrid.FixedCols)) or
     ((rct.top<FGrid.FixedRows) and (rct.Bottom>=FGrid.FixedRows)) then
  begin
    result := 4;
    exit;
  end;

  inc(FCount);
  setlength(FMergeArea, FCount);
  FMergeArea[FCount - 1] := rct;
  result := 0;
end;

//добавить "объединённую" область
//принимает прямоугольник с коорд (x1, y1, x2, y2: integer)
//возвращает:
//      0 - всё нормально, область добавилась
//      1 - указанная область выходит за границы грида, область не добавлена
//      2 - указанная область пересекается(входит) в введённые области , область не добавлена
//      3 - область из одной ячейки не добавляет
function TMergeCells.AddRectXY(x1, y1, x2, y2: integer): byte;
var
  t: TRect;
  
begin
  if x1>x2 then
  begin
    t.Left := x2;
    t.Right := x1;
  end else
  begin
    t.Left := x1;
    t.Right := x2;
  end;
  if y1>y2 then
  begin
    t.Top := y2;
    t.Bottom := y1;
  end else
  begin
    t.Top := y1;
    t.Bottom := y2;
  end;
  result := AddRect(t);
end;

//удаляет облать num
//возвращает
//      - TRUE - область удалена
//      - FALSE - num>Count-1 или num<0
function TMergeCells.DeleteItem(num: integer): boolean;
var
  i: integer;

begin
  if (num>count-1)or(num<0) then
  begin
    result := false;
    exit;
  end;
  for i:= num to Count-2 do
    FMergeArea[i] := FMergeArea[i+1];
  dec(FCount);
  setlength(FMergeArea,FCount);  
  result := true;
end;

//Проверка на вхождение ячейки в левый верхний угол области (ячейка с текстом и "стилем" области)
//принимает (ACol, ARow: integer) -  координаты ячейки
//возвращает
//      -1              ячейка не является левым верхним углом области
//      int num >=0     номер области, в которой ячейка есть левый верхний угол
function TMergeCells.InLeftTopCorner(ACol, ARow: integer): integer;
var
  i, t: integer;

begin
  t := -1;
  for i := 0 to FCount - 1 do
  if (ACol = FMergeArea[i].Left) and (ARow = FMergeArea[i].top) then
  begin
    t:= i;
    break;
  end;
  result := t;
end;

//Проверка на вхождение ячейки в область
//принимает (ACol, ARow: integer) -  координаты ячейки
//возвращает
//      -1              ячейка не входит в область
//      int num >=0     номер области, в которой содержится ячейка
function TMergeCells.InMergeRange(ACol, ARow: integer): integer;
var
  i, t: integer;

begin
  t := -1;
  for i := 0 to FCount - 1 do
  if (ACol >= FMergeArea[i].Left) and (ACol <= FMergeArea[i].Right) and
     (ARow >= FMergeArea[i].Top) and (ARow <= FMergeArea[i].Bottom) then
  begin
    t := i;
    break;
  end;
  result := t;
end;

//Получить ширину области номер num
//возращает
//      -1          - нет такой области
//      int width   - ширина области
function TMergeCells.GetWidthArea(num: integer): integer;
var
  i, t: integer;

begin
  if (num>count-1)or(num<0) then
    t := -1
  else
  begin
    t := 0;
    for i := Items[num].Left to Items[num].Right do
      t := t + FGrid.ColWidths[i];
    if (goFixedVertLine in FGrid.options) or (goVertLine in FGrid.options) then
      t := t + (items[num].Right - items[num].Left) * FGrid.GridLineWidth;
  end;
  result := t;
end;

//Получить высоту области номер num
//возращает
//      -1           - нет такой области
//      int height   - высота области
function TMergeCells.GetHeightArea(num: integer): integer;
var
  i, t: integer;

begin
  if (num>count-1)or(num<0) then
    t := -1
  else
  begin
    t := 0;
    for i := Items[num].top to Items[num].bottom do
      t := t + FGrid.RowHeights[i];
    if (goFixedHorzLine in FGrid.options) or (goHorzLine in FGrid.options) then
      t := t + (Items[num].Bottom - Items[num].Top) * FGrid.GridLineWidth;
  end;    
  result := t;                   {tut}
end;

function TMergeCells.Get_Item(Num: Integer): TRect;
var
  t: TRect;
  
begin
  if (Num>Count-1)or(Num<0) then
  begin
    //а что возвращать в этом случае?
    t.Left := 0;
    t.Right := 0;
    t.Bottom := 0;
    t.Top := 0;
    result := t;
  end else
  result := FMergeArea[Num];
end;

//::: TZColorStringGrid :::

//инициализация
procedure TZColorStringGrid.Initialize();
var
  i: integer;

begin
  FUseBitMap := false;
  UseCellWordWrap := false;
  UseCellSizingHeight := false;
  UseCellSizingWidth := false;
  SizingHeight := false;
  SizingWidth := false;
  WordWrap := false;
  FBmp := TBitmap.Create;
  FBmp.Width := width;
  FBmp.Height := height;
  FDefaultCellParam := TCellStyle.Create(Self);
  FDefaultFixedCellParam := TCellStyle.Create(Self);
  //FDefaultFixedCellParam.BorderCellStyle := sgRaised;
  with FDefaultCellParam do
  begin
    FFont.Assign(Self.Font);
    FBGColor := self.Color;
  end;
  with FDefaultFixedCellParam do
  begin
    FFont.Assign(Self.Font);
    FBGColor := self.FixedColor;
  end;
  setlength(FCellStyle, ColCount);
  for i:= 0 to ColCount - 1 do
  begin
    Setlength(FCellStyle[i], RowCount);
  end;
  FSelectedColors := TSelectColor.Create(self);
  FSelectedColors.FontColor := ClWhite;
  FSelectedColors.BGColor := clHighLight;
  FMergeCells := TMergeCells.Create(self);
  FLineDesign := TLineDesign.Create(self);
  FZInplaceEditor := TZInplaceEditor.Create(self);
  FZInplaceEditor.Visible := false;
  FZInplaceEditor.Top := 1000;
  for i:=0 to 6 do
  begin
    FLastCell[0,i] := -1;
    FLastCell[1,i] := -1;
  end;
  FOptions := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect];
  FInplaceEditorOptions := TInplaceEditorOptions.Create(self);
  DoubleBuffered := true;
  FDontDraw := false;
end;

constructor TZColorStringGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Initialize;
end;

function TZColorStringGrid.GetEditorMode: Boolean;
begin
  if assigned(ZInplaceEditor) then
    result := ZInplaceEditor.Focused
  else
    result := false;
end;

procedure TZColorStringGrid.SetEditorMode(Value: Boolean);
begin
  if value then
  begin
    if (goEditing in Options) and (not (goRowSelect in Options)) then
      ShowEditor
  end else
  begin
    HideEditor;
    if canfocus then setFocus;
  end;
end;

function TZColorStringGrid.GetOptions: TGridOptions;
begin
  result := inherited options;
  if FgoEditing then result := result + [goEditing];
end;

procedure TZColorStringGrid.SetOptions(Value: TGridOptions);
begin
  if goEditing in Value then
    FgoEditing := true
  else
    FgoEditing := false;
  Value := Value - [goEditing];
  inherited options := Value;
end;

//когда получает фокус ввода
procedure TZColorStringGrid.DoEnter;
begin
  {проблема была- грид получал фокус после едитора, поведение не совпадало с StringGrid-ом
  для этого введено FExEn}
  if ZInplaceEditor.FExEn=0 then
  begin
    if (goAlwaysShowEditor in Options) and (goEditing in Options) and (not (goRowSelect in Options)) then
      ShowEditor;
    inherited DoEnter
  end else
    ZInplaceEditor.FExEn := 0;
  invalidate;
end;

//когда контрол теряет фокус
procedure TZColorStringGrid.DoExit;
begin
  if (not ZInplaceEditor.Focused)or(ZInplaceEditor.FExEn=3) then
  begin
    inherited DoExit;
    invalidate;
  end else ZInplaceEditor.FExEn := 1;
end;

//меняет позицию эдитора при скролировании стринг грида
procedure TZColorStringGrid.TopLeftChanged;
begin
  inherited TopLeftChanged;
  if ZInplaceEditor.visible then
  begin
    HideEditor;
    Invalidate;
  end;
end;

procedure TZColorStringGrid.Dblclick;
var
  t: TGridCoord;

begin
  inherited DblClick;
  t := self.MouseCoord(FMouseXY.x, FMouseXY.y);
  if ((t.X = Col) and (t.Y = Row))and(not ZInplaceEditor.Visible) then
    ShowEditorFirst else
  if (ZInplaceEditor.Visible) then
    ShowEditor
  else
    HideEditor;
end;

//(В редакторе) увеличивать размер ячейки при вводе тектса
procedure TZColorStringGrid.SetEditText(ACol, ARow: Longint; const Value: string);
var
  T: TColor;

begin
  t := font.Color;
  font.Assign(CellStyle[ACol, ARow].Font);
  inherited SetEditText(ACol, ARow, Value);
  CalcCellSize(ACol, ARow);
  font.Color := t;
end;

//Рассчитывает параметры для прорисовки текста
//INPUT
//      ACol: integer                 - Колонка ячейки
//      ARow: integer                 - Ряд ячейки
//  var frmparams: integer            - параметры
//  var AlignmentHorizontal: integer  - горизонтальное выравнивание
//  var AlignmentVertical: integer    - вертикальное выравнивание
//  var ZWordWrap: boolean            - перенос слов
procedure TZColorStringGrid.GetParamsForText(ACol, ARow: integer; var frmparams: integer; var AlignmentHorizontal, AlignmentVertical: integer; var ZWordWrap: boolean);
begin
  frmparams := 0;
  ZWordWrap := false;
  AlignmentVertical := 0;
  AlignmentHorizontal := 0;

  case CellStyle[ACol, ARow].HorizontalAlignment of
    taLeftJustify   : AlignmentHorizontal := 0;
    taCenter        : AlignmentHorizontal := 1;
    taRightJustify  : AlignmentHorizontal := 2;
  end;

  case CellStyle[ACol, ARow].VerticalAlignment of
    vaTop   : AlignmentVertical := 0;
    vaCenter: AlignmentVertical := 1;
    vaBottom: AlignmentVertical := 2;
  end;

  if FUseCellWordWrap then
    ZWordWrap := CellStyle[ACol,ARow].WordWrap
  else
    ZWordWrap := FWordWrap;

  if FUseCellSizingHeight then
  begin
    if CellStyle[ACol,ARow].SizingHeight then
      frmparams := frmparams or ZCF_SIZINGH;
  end else
  if FSizingHeight then
    frmparams := frmparams or ZCF_SIZINGH;

  if FUseCellSizingWidth then
  begin
    if CellStyle[ACol,ARow].SizingWidth then
      frmparams := frmparams or ZCF_SIZINGW;
  end else
    if FSizingWidth then
      frmparams := frmparams or ZCF_SIZINGW;
end; //GetParamsForText

//получить размер ячейки
//INPUT
//  var ARect: TRect            - Возвращаемый размер ячейки
//  var frmparams: integer      - возвращаемые параметры для рисования
//      ACol: integer           - Колонка
//      ARow: integer           - Ряд
//  var NumMergerArea: integer  - номер объединённой области
procedure TZColorStringGrid.GetCalcRect(var ARect: TRect; var frmparams: integer; ACol, ARow: integer; var NumMergeArea: integer);
var
  AlignmentHorizontal: integer;
  AlignmentVertical: integer;
  ZWordWrap: boolean;

begin
  NumMergeArea := -1;

  ARect.Top := 0;
  ARect.Left := 0;
  NumMergeArea := MergeCells.InMergeRange(ACol, ARow);
  if (NumMergeArea >= 0) then
  begin
    ARect.Right := MergeCells.GetWidthArea(NumMergeArea);
    ARect.Bottom := MergeCells.GetHeightArea(NumMergeArea);
  end else
  begin
    ARect.Right := self.ColWidths[ACol];
    ARect.Bottom := self.RowHeights[ARow];
  end;

  GetParamsForText(ACol, ARow, frmparams, AlignmentHorizontal, AlignmentVertical, ZWordWrap);

  ZCWriteTextFormatted(Canvas,
                       Cells[ACol, ARow],
                       CellStyle[ACol, ARow].Font,
                       AlignmentHorizontal,
                       AlignmentVertical,
                       ZWordWrap,
                       ARect,
                       CellStyle[ACol, ARow].IndentH,
                       CellStyle[ACol, ARow].IndentV,
                       frmparams or ZCF_CALCONLY,
                       0,                           //Потом можно будет добавить
                       CellStyle[ACol, ARow].Rotate);
end;

//устанавливает размер ячейки
procedure TZColorStringGrid.CalcCellSize(ACol, ARow: integer);
var
  ARect: TRect;
  params: integer;
  h, _x, _y: integer;

begin
  GetCalcRect(ARect, params, ACol, ARow, h);
  if (h >= 0) then
  begin
    _x := MergeCells.GetWidthArea(h);
    _y := MergeCells.GetHeightArea(h);
  end else
  begin
    _x := ColWidths[ACol];
    _y := RowHeights[ARow];
  end;

  if (_x < ARect.Right) then
    ColWidths[ACol] := ARect.Right + ColWidths[ACol] - _x;
  if (_y < ARect.Bottom) then
    RowHeights[ARow] := ARect.Bottom + RowHeights[ARow] - _y;
end; //CalcCellSize

//возвращяет текст из ячейки [ACol, ARow]
function TZColorStringGrid.GetCells(ACol, ARow: Integer): string;
begin
  result := inherited Cells[ACol, ARow];
end;

//установить ячейке [ACol, Arow] текст + вычислить её размер
procedure TZcolorStringGrid.SetCells(ACol, ARow: Integer; const Value: string);
begin
  if value <> inherited Cells[ACol, ARow] then
  begin
    inherited Cells[ACol, ARow] := Value;
    CalcCellSize(ACol, ARow);
    if Assigned(OnSetEditText) then OnSetEditText(self, ACol, ARow, Value);
    if (FLastCell[1,0]=ACol)and(FLastCell[1,1]=ARow) then
      FLastCell[1,0] := -77;
    invalidate;
  end;
end;

//удаление колонки
procedure TZColorStringGrid.DeleteColumn(ACol: integer);
begin
  if (ACOl>ColCount-1)or(ACol<0) then exit;
  ColumnMoved(ACol, ColCount - 1);
  ColCount := ColCount - 1;
end;

//удаление строки
procedure TZColorStringGrid.DeleteRow(ARow: integer);
begin
  if (ARow>RowCount-1)or(ARow<0) then exit;
  RowMoved(ARow, RowCount - 1);
  RowCount := RowCount - 1;
end;

//перетаскивание cтолбца
procedure TZColorStringGrid.ColumnMoved(FromIndex, ToIndex: Longint);
var
  tmp: TCellLine;
  i: integer;

begin
  tmp := FCellStyle[FromIndex];
  if (FromIndex > ToIndex) then
    for i := FromIndex-1 downto ToIndex do
      FCellStyle[i+1] := FCellStyle[i]
  else
    for i := FromIndex to ToIndex-1 do
      FCellStyle[i] := FCellStyle[i+1];
  FCellStyle[ToIndex] := tmp;
  Invalidate;
  inherited ColumnMoved(FromIndex, ToIndex);
end;

//перетаскивание "стиля" строки
procedure TZColorStringGrid.StyleRowMove(FromIndex, ToIndex: integer);
var
  i, j: integer;
  tmp: TCellStyle;

begin
  if FromIndex>ToIndex then
  begin
    for i := 0 to ColCount - 1 do
    begin
      tmp := FCellStyle[i][FromIndex];
      for j:=FromIndex-1 downto ToIndex do
        FCellStyle[i][j+1] := FCellStyle[i][j];
      FCellStyle[i][ToIndex] := tmp;
    end;
  end else
  begin
    for i := 0 to ColCount - 1 do
    begin
      tmp := FCellStyle[i][FromIndex];
      for j:=FromIndex to ToIndex-1 do
        FCellStyle[i][j] := FCellStyle[i][j+1];
      FCellStyle[i][ToIndex] := tmp;
    end;
  end;
end;

//перетаскивание строки
procedure TZColorStringGrid.RowMoved(FromIndex, ToIndex: Longint);
begin
  StyleRowMove(FromIndex, ToIndex);
  inherited RowMoved(FromIndex, ToIndex);
end;

//получить кол-во фиксированных столбцов
function TZColorStringGrid.GetFixedCols(): integer;
begin
  result := inherited FixedCols;
end;

//установить кол-во фиксированных столбцов
procedure TZColorStringGrid.SetFixedCols(const Value: integer);
begin
  inherited FixedCols := Value;
end;

//получить кол-во фиксированных строк
function TZColorStringGrid.GetFixedRows(): integer;
begin
  result := inherited FixedRows;
end;

//установить кол-во фиксированных строк
procedure TZColorStringGrid.SetFixedRows(const Value: integer);
begin
  inherited FixedRows := Value;
end;

//получить цвет фиксированной ячейки
function TZColorStringGrid.GetFixedColor: TColor;
begin
  result := inherited FixedColor;
end;

//установить цвет фиксированной ячейки
procedure TZColorStringGrid.SetFixedColor(const Value: TColor);
begin
  inherited FixedColor := value;
  FDefaultFixedCellParam.BGColor := value;
  Invalidate;
end;

//изменить кол-во строк/столбцов
procedure TZColorStringGrid.SetCellCount(oldColCount, newColCount, oldRowCount, newRowCount: integer);
var
  i, j: integer;

begin
  //столбцы
  if oldColCount>newColCount then
  begin
    for i:= oldColCount-1 to newColCount-1  do
    begin
      for j:=0 to newRowCount do
        FCellStyle[i][j].Free;
      Setlength(FCellStyle[i], 0);
    end;
    setlength(FCellStyle, newcolCount);
  end else
  if oldColCount<newColCount then
  begin
    setlength(FCellStyle, newcolCount);
    for i:= oldColCount-1 to newColCount-1  do
    begin
      Setlength(FCellStyle[i], newRowCount);
    end;
  end;
  //строки
  if oldRowCount>newRowCount then
  begin
    for i:=0 to newColCount-1 do
    begin
      for j:= newRowCount to oldRowCount-1 do
      begin
        if FCEllStyle[i][j] <> nil then
        begin
          FCellStyle[i][j].Free;
        end;
      end;
      Setlength(FCellStyle[i], newRowCount);
    end;
  end else
  if oldRowCount<newRowCount then
  begin
    for i:=0 to newColCount-1 do
      Setlength(FCellStyle[i], newRowCount);
  end;
end;

//установить количество строк
procedure TZColorStringGrid.SetRowCount(const Value: integer);
begin
  SetCellCount(ColCount, ColCount, inherited RowCount, value);
  inherited RowCount := Value;
end;

//прочистать количесвто строк
function TZColorStringGrid.GetRowCount: integer;
begin
  result := inherited RowCount;
end;

//установить количество столбцов
procedure TZColorStringGrid.SetColCount(const Value: integer);
begin
  SetCellCount(inherited ColCount, value, RowCount, RowCount);
  inherited ColCount := Value;
end;

//прочистать количесвто столбцов
function TZColorStringGrid.GetColCount: integer;
begin
  result := inherited ColCount;
end;

//установить оформление ячеек по умолчанию
procedure TZColorStringGrid.SetDefaultCellParam(const Value: TCellStyle);
begin
  FDefaultCellParam.Assign(Value);
  self.Color := value.BGColor;
  Invalidate;
end;

//прочитать оформление ячеек
function TZColorStringGrid.GetDefaultCellParam: TCellStyle;
begin
  result := FDefaultCellParam;
end;

//установить оформление фиксированных ячеек по умолчанию
procedure TZColorStringGrid.SetDefaultFixedCellParam(const Value: TCellStyle);
begin
  FDefaultFixedCellParam.Assign(Value);
  FixedColor := value.FBGColor;
  Invalidate;
end;

//прочитать оформление фиксированных ячеек
function TZColorStringGrid.GetDefaultFixedCellParam: TCellStyle;
begin
  result := FDefaultFixedCellParam;
end;

//установить оформление ячейки [ACol, ARow]
procedure TZColorStringGrid.SetCellStyle(ACol, ARow: integer; const Value: TCellStyle);
begin
  if (ACol>ColCount-1) or (ARow>RowCount-1) or (ACOl<0) or (ARow<0) then exit;
  if FCellStyle[ACol][ARow] = nil then
    FCellStyle[ACol][ARow] := TCellStyle.Create(self);
  FCellStyle[ACol][ARow].Assign(Value);
end;

//прочитать оформление ячейки [ACol, ARow]
function TZColorStringGrid.GetCellStyle(ACol, ARow: integer): TCellStyle;
begin
  if (ACol>ColCount-1) or (ARow>RowCount-1) or (ACOl<0) or (ARow<0) then
  begin
    result := FDefaultCellParam;
    exit;
  end;
  if FCellStyle[ACol][ARow]=nil then
  begin
    FCellStyle[ACol][ARow] := TCellStyle.Create(self);
    if (ACol < FixedCols)or(ARow < FixedRows) then
      FCellStyle[ACol][ARow].Assign(DefaultFixedCellStyle)
    else
      FCellStyle[ACol][ARow].assign(DefaultCellStyle);
  end;
  result := FCellStyle[ACol][ARow];
end;

//устанавливает стиль у всей строки
//   ARow - номер строки
//   fixedCol - менять ли  стиль у фиксированных ячеек (true - да)
procedure TZColorStringGrid.SetCellStyleRow(ARow: integer; fixedCol: boolean; const value: TCellStyle);
var
  i, t: integer;

begin
  if (ARow>RowCount-1) or (ARow<0) then exit;
  if fixedCol then t := 0 else t := FixedCols;
  for i := t to ColCount - 1 do
  begin
    if FCellStyle[i][ARow] = nil then
      FCellStyle[i][ARow] := TCellStyle.Create(self);
    FCellStyle[i][ARow].Assign(Value);
  end;
  invalidate;
end;

//устанавливает стиль у столбца
//   ACol - номер столбца
//   fixedRow - менять ли  стиль у фиксированных ячеек (true - да)
procedure TZColorStringGrid.SetCellStyleCol(ACol: integer; fixedRow: boolean; const value: TCellStyle);
var
  i, t: integer;

begin
  if (ACol>ColCount-1) or (ACol<0) then exit;
  if fixedRow then t := 0 else t := FixedRows;
  for i := t to RowCount - 1 do
  begin
    if FCellStyle[ACol][i] = nil then
      FCellStyle[ACol][i] := TCellStyle.Create(self);
    FCellStyle[ACol][i].Assign(Value);
  end;
  invalidate;
end;

function TZColorStringGrid.DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint): Boolean;
var
  t: integer;

begin
  t := MergeCells.InMergeRange(Col, row);
  Result := inherited DoMouseWheelDown(Shift, MousePos);
  if t>=0 then
  begin
    if MergeCells.Items[t].Bottom<RowCount -2 then
    begin
      Result := true;
      row := MergeCells.Items[t].Bottom + 1;
    end else Result := false;
  end;
  RowSelectYY(VK_DOWN);
end;

function TZColorStringGrid.DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint): Boolean;
var
  t: integer;

begin
  t := MergeCells.InMergeRange(Col, Row);
  Result := inherited DoMouseWheelUp(Shift, MousePos);
  if t>=0 then
  begin
    if MergeCells.Items[t].Top>FixedRows then
    begin
      Result := true;
      row := MergeCells.Items[t].Top - 1;
    end else Result := false;
  end;
  RowSelectYY(VK_UP);
end;

//Выделение строки при goRowSelect in Options при объеденённых ячейках 
procedure TZColorStringGrid.RowSelectYY(key: word);
var
  z: TPoint;
  k: TGridRect;

begin
  if goRowSelect in Options then
  begin
    z := MergeCells.GetMergeYY(Row);
    k:=selection;
    if not(((z.y = k.Top) and (z.x = k.Bottom)) or ((z.x = k.Top) and (z.y = k.Bottom))) then
    case key of
      VK_UP:
        begin
          k.Top := z.y;
          k.Bottom := z.x;
        end;
      VK_DOWN:
        begin
          k.Top := z.x;
          k.Bottom := z.y;
        end;
    end;
    selection := k;
  end;
end;

procedure TZColorStringGrid.SavePenBrush(var CellCanvas: TCanvas);
begin
  with CellCanvas do
  begin
    FPenBrush.PColor := Pen.Color;
    FPenBrush.Width := Pen.Width;
    FPenBrush.Mode := Pen.Mode;
    FPenBrush.PStyle := Pen.Style;
    FPenBrush.BColor := Brush.Color;
    FPenBrush.BStyle := Brush.Style;
  end;
end;

Procedure TZColorStringGrid.RestorePenBrush(var CellCanvas: TCanvas);
begin
  with CellCanvas do
  begin
    Pen.Color := FPenBrush.PColor;
    Pen.Width := FPenBrush.Width;
    Pen.Mode := FPenBrush.Mode;
    Pen.Style := FPenBrush.PStyle;
    Brush.Color := FPenBrush.BColor;
    Brush.Style := FPenBrush.BStyle;
  end;
end;

procedure TZColorStringGrid.KeyDown(var Key: Word; Shift: TShiftState);
var
  t: integer;
  k, l: TGridRect;

  function test_match: boolean;
  begin
    k := self.MergeCells.GetSelectedArea(false);
    if goRowSelect in Options then
    begin
      k.Left := self.FixedCols;
      k.Right := self.ColCount - 1;
    end;
    result := (selection.Left <> k.Left)or(selection.right <> k.right) or
              (selection.top <> k.top)or(selection.bottom <> k.bottom);
  end;

begin
  t := MergeCells.InMergeRange(Col, row);
  inherited KeyDown(Key, Shift);
  if ssShift in shift then
  begin
    case key of
      VK_UP:
        begin
          if test_match then
          begin
            l := k;
            k.Top := l.Bottom;
            k.Bottom := l.Top;
            Selection := k;
          end;
          if goRowSelect in Options then
            Invalidate;
        end;
      VK_DOWN:
        begin
          if test_match then
            Selection := k;
          if goRowSelect in Options then
            Invalidate;
        end;
      VK_LEFT:
        if test_match then
        begin
          l := k;
          k.left := l.right;
          k.right := l.left;
          Selection := k;
        end;
      VK_RIGHT:
        if test_match then
          Selection := k;
      VK_F2:
      begin
        ShowEditorFirst;
        ZInplaceEditor.SelStart := length(Cells[Col, Row]);
      end;
    end;
  end else
  case key of
    VK_UP:
      begin
        if row>FixedRows then
        begin
          if t>=0 then
          if MergeCells.Items[t].Top > FixedRows then
            row := MergeCells.Items[t].Top-1
          else
            row := MergeCells.Items[t].Top;
        end;
        RowSelectYY(VK_UP);
        key := 0;
      end;
    VK_DOWN:
      begin
        if row<RowCount-1 then
        begin
          if t>=0 then
          if MergeCells.Items[t].Bottom + 1 = self.RowCount then
            row := MergeCells.Items[t].Bottom
          else
            row := MergeCells.Items[t].Bottom + 1;
        end;
        RowSelectYY(VK_DOWN);
        key := 0;
      end;
    VK_LEFT:
      begin
        if col>FixedCols then
        begin
          if t>=0 then
          if MergeCells.Items[t].left > FixedCols then
            col := MergeCells.Items[t].left-1
          else
            col := MergeCells.Items[t].left;
        end;
        RowSelectYY(VK_UP);
      end;
    VK_RIGHT:
      begin
        if col<ColCount-1 then
        begin
          if t>=0 then
          if MergeCells.Items[t].Right + 1 = ColCount then
            col := MergeCells.Items[t].Right
          else
            col := MergeCells.Items[t].Right + 1;
        end;
        RowSelectYY(VK_DOWN);
      end;
    VK_NEXT:
      RowSelectYY(VK_DOWN);
    VK_PRIOR:
      RowSelectYY(VK_UP);
    VK_F2:
    begin
      ShowEditorFirst;
      ZInplaceEditor.SelStart := length(Cells[Col, Row]);
    end;
  end;
end;

procedure TZColorStringGrid.KeyPress(var Key: Char);
begin
  if assigned(onKeyPress) then onKeyPress(self, key);
  if (not EditorMode)and ((goEditing in Options)and(not (goRowSelect in Options))) then
  begin
    if key = #13 then
    begin
      ShowEditorFirst;
      if ZInplaceEditor.Visible then
        ZInplaceEditor.SelectAll;
    end else
    if ord(key) > 32 then
    begin
      ShowEditorFirst;
      ZInplaceEditor.Lines.Text := key;
      ZInplaceEditor.SelStart := 1;
    end;
  end;
end;

procedure TZColorStringGrid.Paint;
var
  DrawInfo: TGridDrawInfo;
  Sel: TGridRect;
  TmpRect, r1, r2: TRect;
  i, j: integer;

  function _isEmptyCell(const rr: TRect): boolean;
  begin
    result := (rr.Left = rr.Right) and (rr.Top = rr.Bottom);
  end;

  procedure DrawLines(var Rct: Trect; _x,_y: integer);
  begin
    if GridLineWidth>0 then
    begin
      canvas.brush.Color := LineDesign.LineColor;
      if (goFixedHorzLine in options) or (goHorzLine in options) then
      begin
        TmpRect.Left   := rct.left;
        TmpRect.Right  := rct.left + DrawInfo.Horz.EffectiveLineWidth + _x;
        TmpRect.Top    := rct.top+ _y;
        TmpRect.Bottom := rct.top + DrawInfo.Vert.EffectiveLineWidth + _y;
        canvas.FillRect(TmpRect);
      end;
      if (goFixedVertLine in options) or (goVertLine in options) then
      begin
        TmpRect.Left   := rct.left + _x;
        TmpRect.Right  := rct.left + DrawInfo.Horz.EffectiveLineWidth + _x;
        TmpRect.Top    := rct.top;
        TmpRect.Bottom := rct.top + DrawInfo.Vert.EffectiveLineWidth + _y;
        canvas.FillRect(TmpRect);
      end;
    end;
  end;

  procedure drawRect(var where: TRect; colorUp, colorDown: TColor);
  begin
    canvas.Pen.Color := clred;
    canvas.pen.Color := colorDown;
    canvas.pen.Width := 1;
    canvas.MoveTo(where.left, where.Top);
    canvas.LineTo(where.left, where.bottom);
    canvas.MoveTo(where.left, where.Top);
    canvas.LineTo(where.right, where.Top);
    canvas.pen.Color := colorUp;
    canvas.MoveTo(where.right-1, where.Top + 1);
    canvas.LineTo(where.right-1, where.bottom);
    canvas.MoveTo(where.left+1, where.bottom- 1);
    canvas.LineTo(where.right, where.bottom - 1);
  end;

  procedure DrawCells(ACol, ARow: Longint; StartX, StartY, StopX, StopY: Integer;
    Color: TColor; IncludeDrawState: TGridDrawState);
  var
    CurCol, CurRow: Longint;
    Where, MCI: TRect;
    DrawState: TGridDrawState;
    t, _x, _y: integer;

  begin
    CurRow := ARow;
    Where.Top := StartY;

    while (Where.Top <= StopY) and (CurRow < RowCount) do
    begin
      CurCol := ACol;
      Where.Left := StartX;
      Where.Bottom := Where.Top + RowHeights[CurRow];
      while (Where.Left <= StopX) and (CurCol < ColCount) do
      begin
        Where.Right := Where.Left + ColWidths[CurCol];
        if (Where.Right >= Where.Left) then
        begin
          DrawState := IncludeDrawState;
          t := MergeCells.inMergeRange(CurCol, CurRow);
          //если ячейка не принадлежит объединённой области
          if (t < 0) then
          begin
            DrawCell(CurCol, CurRow, Where, DrawState);
            if (GridLineWidth > 0) then
            begin
              canvas.brush.Color := LineDesign.LineColor;
              if (goFixedHorzLine in options) or (goHorzLine in options) then
              begin
                TmpRect.Left   := where.left;
                TmpRect.Right  := where.right + DrawInfo.Horz.EffectiveLineWidth ;
                TmpRect.Top    := where.Bottom;
                TmpRect.Bottom := where.Bottom + DrawInfo.Vert.EffectiveLineWidth ;
                canvas.FillRect(TmpRect);
              end;
              if (goFixedVertLine in options) or (goVertLine in options) then
              begin
                TmpRect.Left   := where.right;
                TmpRect.Right  := where.right + DrawInfo.Horz.EffectiveLineWidth ;
                TmpRect.Top    := where.top;
                TmpRect.Bottom := where.Bottom + DrawInfo.Vert.EffectiveLineWidth ;
                canvas.FillRect(TmpRect);
              end;
            end;
            {begin стиль ячейки (утопленный/плоский/аля-кнопка)}
            if (csDesigning in ComponentState) then
            begin
              if (ACol < FixedCols)or(ARow < FixedRows) then
              begin
                if (DefaultFixedCellStyle.BorderCellStyle <> sgNone) then
                begin
                  if (DefaultFixedCellStyle.BorderCellStyle = sgRaised) then
                    drawRect(where,LineDesign.LineDownColor, LineDesign.LineUpColor)
                  else
                    drawRect(where,LineDesign.LineUpColor, LineDesign.LineDownColor);
                end;
              end else
              begin
                if (DefaultCellStyle.BorderCellStyle <> sgNone) then
                begin
                  if (DefaultCellStyle.BorderCellStyle = sgRaised) then
                    drawRect(where,LineDesign.LineDownColor, LineDesign.LineUpColor)
                  else
                    drawRect(where,LineDesign.LineUpColor, LineDesign.LineDownColor);
                end;
              end;
            end else
            if (CellStyle[CurCol, CurRow].BorderCellStyle <> sgNone) then
            begin
              if (CellStyle[CurCol, CurRow].BorderCellStyle = sgRaised) then
                drawRect(where,LineDesign.LineDownColor, LineDesign.LineUpColor)
              else
                drawRect(where,LineDesign.LineUpColor, LineDesign.LineDownColor);
            end;
            {end стиль ячейки (утопленный/плоский/аля-кнопка)}
          end else
          begin
            _y := 0;
            MCI := MergeCells.Items[t];
            for _x := MCI.left to CurCol-1 do
              _y := _y + self.ColWidths[_x] + DrawInfo.Horz.EffectiveLineWidth;
            FRctRectangle.Left := _y;
            FRctRectangle.Right := FRctRectangle.Left + ColWidths[CurCol] + DrawInfo.Horz.EffectiveLineWidth;

            _y := 0;
            for _x := MCI.top to CurRow-1 do
            _y := _y + self.RowHeights[_x] + DrawInfo.Vert.EffectiveLineWidth;
            FRctRectangle.top := _y;
            FRctRectangle.bottom := FRctRectangle.top + RowHeights[CurRow] + DrawInfo.vert.EffectiveLineWidth;

            //если ячейка содержиться в области для объединения
            //рисуем всю область
            TmpRect := CellRect(MCI.left, MCI.top);
            FUseBitMap := true;
            FRctBitMap := where;
            FRctBitMap.right := FRctBitMap.right + DrawInfo.Horz.EffectiveLineWidth;
            FRctBitMap.bottom := FRctBitMap.bottom + DrawInfo.Vert.EffectiveLineWidth;
            //какая ячейка прорисовывается
            FLastCell[0,0] := MCI.left;
            FLastCell[0,1] := MCI.top;
            FLastCell[0,2] := TmpRect.Right - TmpRect.Left;
            FLastCell[0,3] := TmpRect.Bottom - TmpRect.Top;
            FLastCell[0,5] := FRctRectangle.Bottom - FRctRectangle.top;
            FLastCell[0,6] := FRctRectangle.Right - FRctRectangle.Left;

            if ((Selection.Top >= MCI.Top) and
               (selection.Left >= MCI.left) and
               (Selection.Top <= MCI.bottom) and
               (Selection.Left <= MCI.right))or
               ((MCI.Top >= Selection.Top) and   // Если объединённая ячейка
               (MCI.Top <= Selection.Bottom) and // в выделенной области
               (MCI.Left >= Selection.Left) and
               (MCI.Left <= Selection.Right)) then
            begin
              FLastCell[0,4] := 1;
              DrawCell(MCI.left, MCI.top, TmpRect, DrawState + [gdSelected]);
            end else
            begin
              FLastCell[0,4] := 0;
              DrawCell(MCI.left, MCI.top, TmpRect, DrawState);
            end;

            //запоминаем, какую ячейку рисовали
            FLastCell[1,0] := FLastCell[0,0];
            FLastCell[1,1] := FLastCell[0,1];
            FLastCell[1,2] := FLastCell[0,2];
            FLastCell[1,3] := FLastCell[0,3];
            FLastCell[1,4] := FLastCell[0,4];
            FLastCell[1,5] := FLastCell[0,5];
            FLastCell[1,6] := FLastCell[0,6];
            FLastCell[0,0] := - 10;

            FUseBitMap := false;
            //рисуем границы

            // right
            if (CurCol = MCI.Right) then
            begin
              if (GridLineWidth > 0) then
              begin
                if (goFixedVertLine in options) or (goVertLine in options) then
                begin
                  canvas.brush.Color := LineDesign.LineColor;
                  TmpRect.Left   := where.right;
                  TmpRect.Right  := where.right + DrawInfo.Horz.EffectiveLineWidth ;
                  TmpRect.Top    := where.top;
                  TmpRect.Bottom := where.Bottom + DrawInfo.Vert.EffectiveLineWidth ;
                  canvas.FillRect(TmpRect);
                end;
              end;

              if (CellStyle[MCI.Left, MCI.top].BorderCellStyle <> sgNone) then
              begin
                canvas.pen.Width := 1;
                if (CellStyle[MCI.Left, MCI.top].BorderCellStyle = sgRaised) then
                  canvas.pen.Color := LineDesign.LineDownColor
                else
                  canvas.pen.Color := LineDesign.LineUpColor;
                canvas.MoveTo(where.right-1, where.Top);
                canvas.LineTo(where.right-1, where.bottom + GridLineWidth);
              end;
            end; //right

            //left
            if CurCol = MCI.left then
            begin
              if (CellStyle[MCI.Left, MCI.top].BorderCellStyle <> sgNone) then
              begin
                if (CellStyle[MCI.Left, MCI.top].BorderCellStyle = sgRaised) then
                  canvas.pen.Color := LineDesign.LineUpColor
                else
                  canvas.pen.Color := LineDesign.LineDownColor;
                canvas.pen.Width := 1;
                if (CurRow = MCI.top) then
                  canvas.MoveTo(where.left, where.Top)
                else
                  canvas.MoveTo(where.left, where.Top-GridLineWidth);
                canvas.LineTo(where.left, where.bottom);
              end;
            end; //left

            // bottom
            if (CurRow = MCI.Bottom) then
            begin
              if (GridLineWidth > 0) then
              begin
                if (goFixedHorzLine in options) or (goHorzLine in options) then
                begin
                  canvas.brush.Color := LineDesign.LineColor;
                  TmpRect.Left   := where.left;
                  TmpRect.Right  := where.right + DrawInfo.Horz.EffectiveLineWidth;
                  TmpRect.Top    := where.Bottom;
                  TmpRect.Bottom := where.Bottom + DrawInfo.Vert.EffectiveLineWidth;
                  canvas.FillRect(TmpRect);
                end;
              end;

              if (CellStyle[MCI.Left, MCI.top].BorderCellStyle <> sgNone) then
              begin
                if (CellStyle[MCI.Left, MCI.top].BorderCellStyle = sgRaised) then
                  canvas.pen.Color := LineDesign.LineDownColor
                else
                  canvas.pen.Color := LineDesign.LineUpColor;

                canvas.pen.Width := 1;
                canvas.MoveTo(where.left, where.bottom - 1);
                if (CurCol <> MergeCells.Items[t].Right) then
                   canvas.LineTo(where.right + GridLineWidth, where.bottom - 1)
                else
                  canvas.LineTo(where.right, where.bottom - 1);
              end;
            end; //bottom

            //top
            if (CurRow = MCI.top) then
            begin
              if (CellStyle[MCI.Left, MCI.top].BorderCellStyle <> sgNone) then
              begin
                if (CellStyle[MCI.Left, MCI.Top].BorderCellStyle = sgRaised) then
                  canvas.pen.Color := LineDesign.LineUpColor
                else
                  canvas.pen.Color := LineDesign.LineDownColor;
                canvas.pen.Width := 1;
                canvas.MoveTo(where.left, where.Top);
                if (CurCol <> MCI.Right) then
                  canvas.LineTo(where.right + GridLineWidth, where.top)
                else
                  canvas.LineTo(where.right, where.top);
              end;
            end; //top

          end;
        end;
        Where.Left := Where.Right + DrawInfo.Horz.EffectiveLineWidth;
        Inc(CurCol);
      end; //while
      Where.Top := Where.Bottom + DrawInfo.Vert.EffectiveLineWidth;
      Inc(CurRow);
    end; //while
  end;

begin
  if (FDontDraw) then
    exit;

  if (ZInplaceEditor.visible)or((goAlwaysShowEditor in Options)and(goEditing in Options)and(not(goRowSelect in Options))) then
    showeditor;
  CalcDrawInfo(DrawInfo);
  with DrawInfo do
  begin
    DrawCells(0, 0, 0, 0, Horz.FixedBoundary, Vert.FixedBoundary, FixedColor,
      [gdFixed]);
    DrawCells(LeftCol, 0, Horz.FixedBoundary, 0, Horz.GridBoundary,
      Vert.FixedBoundary, FixedColor, [gdFixed]);
    DrawCells(0, TopRow, 0, Vert.FixedBoundary, {FixedRows + 10}Horz.FixedBoundary,
      Vert.GridBoundary, FixedColor, [gdFixed]); {?}
    DrawCells(LeftCol, TopRow, Horz.FixedBoundary,
      Vert.FixedBoundary, Horz.GridBoundary, Vert.GridBoundary, Color, []);

    // фокус на ячейке
    if goRowSelect in Options then
    begin
      if focused then
      if SelectedColors.UseFocusRect then
      begin
        tmpRect := BoxRect(selection.Left, Selection.Top, Selection.Right, Selection.Bottom);
        canvas.DrawFocusRect(tmpRect);
      end;
    end else
    if not((goAlwaysShowEditor in Options)and(goEditing in Options)) then
    begin
      if focused then
      if SelectedColors.UseFocusRect then
      begin
         sel.Top := self.MergeCells.InMergeRange(Col, Row);
         if sel.Top >= 0 then
         begin
           //что-то BoxRect немного не подходит для объединённых ячеек... T_T
           r1 := CellRect(MergeCells.Items[sel.top].Left, MergeCells.Items[sel.top].top);
           r2 := CellRect(MergeCells.Items[sel.top].Right, MergeCells.Items[sel.top].Bottom);
           if (not (_isEmptyCell(r1) or _isEmptyCell(r2))) then
           begin
             tmpRect.Left := CellRect(MergeCells.Items[sel.top].Left, MergeCells.Items[sel.top].top).Left;
             tmpRect.Top := CellRect(MergeCells.Items[sel.top].Left, MergeCells.Items[sel.top].top).Top;
             tmpRect.Right := CellRect(MergeCells.Items[sel.top].Right, MergeCells.Items[sel.top].Bottom).right;
             tmpRect.Bottom := CellRect(MergeCells.Items[sel.top].Right, MergeCells.Items[sel.top].Bottom).Bottom;
           end else
           begin
             tmpRect.Left := -1;
             tmpRect.Top := -1;
             tmpRect.Right := -1;
             tmpRect.Bottom := -1;
             for i := MergeCells.Items[sel.top].Left to MergeCells.Items[sel.top].Right do
             for j := MergeCells.Items[sel.top].top to MergeCells.Items[sel.top].Bottom do
             begin
               r1 := CellRect(i, j);
               if (not _isEmptyCell(r1)) then
               begin
                 if (tmpRect.Left < 0) or (tmpRect.Left > r1.Left) then
                   tmpRect.Left := r1.Left;
                 if (tmpRect.Top < 0) or (tmpRect.Top > r1.Top) then
                   tmpRect.Top := r1.Top;
                 if (tmpRect.Left < 0) or (tmpRect.Bottom < r1.Bottom) then
                   tmpRect.Bottom := r1.Bottom;
                 if (tmpRect.Right < 0) or (tmpRect.Right < r1.Right) then
                   tmpRect.Right := r1.Right;
               end;
             end;
           end;
         end else
         begin
           tmpRect := BoxRect(Col, Row, Col, Row);
         end;
         canvas.DrawFocusRect(tmpRect);
      end;
    end;

    if Horz.GridBoundary < Horz.GridExtent then
    begin
      Canvas.Brush.Color := Color;
      Canvas.FillRect(Rect(Horz.GridBoundary, 0, Horz.GridExtent, Vert.GridBoundary));
    end;
    if Vert.GridBoundary < Vert.GridExtent then
    begin
      Canvas.Brush.Color := Color;
      Canvas.FillRect(Rect(0, Vert.GridBoundary, Horz.GridExtent, Vert.GridExtent));
    end;
  end;
end;

//отрисовка ячейки
procedure TZColorStringGrid.DrawCell(ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
var
  t: integer;
  zCanvas: TCanvas;
  dx, dy: integer;
  AlignmentHorizontal: integer;
  AlignmentVertical: integer;
  frmparams: integer;
  ZWordWrap: boolean;
  fntColor: TColor;

begin
  if (ZInplaceEditor.Visible) then
    if (ACol = Col)and(ARow = Row) then exit;

  if FUseBitMap then
  begin
    if height > FBmp.Height then FBmp.Height := height;
    if width > FBmp.Width then FBmp.Width := width;
    zCanvas := FBmp.Canvas;
    dx := abs(Arect.Left);
    Arect.Left := 0;
    Arect.right := Arect.Right - dx;
    dy := abs(Arect.top);
    Arect.top := 0;
    Arect.bottom := Arect.bottom - dy;

    t := MergeCells.InLeftTopCorner(ACol, ARow);
    if t>=0 then
    begin
      ARect.Right := ARect.Left + MergeCells.GetWidthArea(t);
      ARect.Bottom := ARect.top + MergeCells.GetHeightArea(t);
      if FUseBitmap then
      begin
        if MergeCells.GetHeightArea(t) > FBmp.Height then FBmp.Height := MergeCells.GetHeightArea(t);
        if MergeCells.GetWidthArea(t) > FBmp.Width then FBmp.Width := MergeCells.GetWidthArea(t);
      end;
    end;

  end else
    zCanvas := Canvas;

  //если текущая ячейка принадлежит той же области, что и рисовавщаяся до этого,
  //то её изображение берём из буфера (FBmp), иначе прорисовываем
  if (not ((FLastCell[0,0]=FLastCell[1,0])and
           (FLastCell[0,1]=FLastCell[1,1])and
           (FLastCell[0,2]=FLastCell[1,2])and
           (FLastCell[0,3]=FLastCell[1,3])and
           (FLastCell[0,4]=FLastCell[1,4])and
           (FLastCell[0,5]=FLastCell[1,5])and
           (FLastCell[0,6]=FLastCell[1,6])))
      or(not FUseBitMap) then
  with zCanvas do
  begin
    if csDesigning in ComponentState then
    begin
      if (ACol < FixedCols)or(ARow < FixedRows) then
        Brush.Color := FDefaultFixedCellParam.BGColor
      else
      begin
        if (goRowSelect in Options) and (ARow=FixedRows)and(FSelectedColors.ColoredSelect) then
          Brush.Color := FSelectedColors.BGColor else
        if (ACol=FixedCols)and(ARow=FixedRows)and(FSelectedColors.ColoredSelect) then
          Brush.Color := FSelectedColors.BGColor
        else
          Brush.Color := FDefaultCellParam.BGColor;
      end;
      FillRect(ARect);
    end else
    begin
      t := MergeCells.InLeftTopCorner(ACol, ARow);
      if t>=0 then
      begin
        {ARect.Right := ARect.Left + MergeCells.GetWidthArea(t);
        ARect.Bottom := ARect.top + MergeCells.GetHeightArea(t);
        if FUseBitmap then
        begin
          if MergeCells.GetHeightArea(t) > FBmp.Height then FBmp.Height := MergeCells.GetHeightArea(t);
          if MergeCells.GetWidthArea(t) > FBmp.Width then FBmp.Width := MergeCells.GetWidthArea(t);
        end; }
      end else
      begin
        t := MergeCells.InMergeRange(ACol, ARow);
        if t >= 0 then exit;
      end;

      fntColor := CellStyle[ACol, ARow].Font.Color;

      if (FSelectedColors.ColoredSelect)and(not (goAlwaysShowEditor in Options)) then
      begin
        if ((ARow >= selection.Top) and
           (ACol >= selection.left) and
           (ARow <= selection.bottom) and
           (ACol <= selection.right))or(gdSelected in Astate) then
        begin
          Brush.Color := SelectedColors.BGColor;
          fntColor := SelectedColors.FontColor;
        end else
          Brush.Color :=CellStyle[ACol,ARow].BGColor;
      end else
        Brush.Color :=CellStyle[ACol,ARow].BGColor;

      FillRect(ARect);

      if FUseBitMap then
      begin
        if Assigned(FOnBeforeTextDrawCellMerge) then
        begin
          //сохраним некоторые свойства pen и brush...
          SavePenBrush(zCanvas);
          FOnBeforeTextDrawCellMerge(self, ACol, ARow, ARect, astate, zCanvas);
          //восстановим pen и brush
          RestorePenBrush(zCanvas);
        end;
        zCanvas.Font.Assign(CellStyle[ACol, ARow].Font);
        if (((ARow >= selection.Top) and
           (ACol >= selection.left) and
           (ARow <= selection.bottom) and
           (ACol <= selection.right)) or (gdSelected in Astate)) and (SelectedColors.ColoredSelect) then
          Font.Color  := SelectedColors.FontColor;
      end else
      begin
        if Assigned(FOnbeforeTextDrawCell) then
        begin
          //сохраним некоторые свойства pen и brush...
          SavePenBrush(zCanvas);
          pen.Color := clblack;
          FOnbeforeTextDrawCell(self, ACol, Arow, ARect, astate);
          //восстановим pen и brush
          RestorePenBrush(zCanvas);
        end;
      end;

      //h := drawText(handle, pchar(Cells[ACol, ARow]), length(Cells[ACol, ARow]), ARect, uFrmt);

      GetParamsForText(ACol, ARow, frmparams, AlignmentHorizontal, AlignmentVertical, ZWordWrap);

      ZCWriteTextFormatted(zCanvas,
                       Cells[ACol, ARow],
                       CellStyle[ACol, ARow].Font,
                       fntColor,
                       AlignmentHorizontal,
                       AlignmentVertical,
                       ZWordWrap,
                       ARect,
                       CellStyle[ACol, ARow].IndentH,
                       CellStyle[ACol, ARow].IndentV,
                       frmparams or ZCF_NO_CLIP,
                       0,                           //Потом можно будет добавить
                       CellStyle[ACol, ARow].Rotate);
      Pen.Width := 1;
    end;
  end;

  DefaultDrawing := false;

  if FUseBitMap then
  begin
    if (ARect.Right-ARect.Left <> 0) then
    if Assigned(FOnDrawCellMerge) then
      FOnDrawCellMerge(self, ACol, ARow, ARect, astate, zCanvas);
    //canvas.CopyRect(FRctBitMap, zCanvas, FRctRectangle);
    BitBlt(Canvas.Handle,
           FRctBitMap.Left,
           FRctBitMap.Top,
           FRctBitMap.Right - FRctBitMap.Left,
           FRctBitMap.Bottom - FRctBitMap.Top,
           ZCanvas.Handle,
           FRctRectangle.Left,
           FRctRectangle.Top,
           SRCCOPY);

  end;
  if not FUseBitMap then
  if Assigned(OnDrawCell) then
    OnDrawCell(self, ACol, ARow, ARect, AState)
end;

procedure TZColorStringGrid.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  t: TGridCoord;
  t2: TGridRect;

begin
  if (goEditing in Options)and(not (goRowSelect in Options))and (goAlwaysShowEditor in Options) then
  begin
    t := MouseCoord(x, y);
    if (Col <> t.y) or (Row <> t.x) then ShowEditorfirst;
  end;
  if ssShift in Shift then
  begin
    t := MouseCoord(x, y);
    t2.Left := t.X;
    t2.top := t.Y;
    if t2.Left < self.FixedCols then t2.Left := self.FixedCols;
    if t2.top < self.FixedRows then t2.Top := self.FixedRows;
    t2.Right := col;
    t2.Bottom := row;
    selection := t2;
    t2 := MergeCells.GetSelectedArea(false);
    selection := t2;
    invalidate;
    if assigned(onMouseDown) then onMouseDown(self, Button, Shift, X, Y);
  end else
    inherited MouseDown(Button, Shift, X, Y);
  FMouseXY.X := x;
  FMouseXY.y := y;
  RowSelectYY(VK_DOWN);
  if (ZInplaceEditor.Visible) or (inherited EditorMode) then
    ShowEditor;
end;

procedure TZColorStringGrid.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  oldsel, sel: TGridRect;
  t, t1 : TGridCoord;

begin
  oldsel := selection;
  inherited;
  //учитываем объединённые ячейки
  sel := selection;
  if (oldsel.Left <> sel.Left) or
     (oldsel.top <> sel.top) or
     (oldsel.right <> sel.right) or
     (oldsel.bottom <> sel.bottom) then
  begin
    oldsel := self.MergeCells.GetSelectedArea(false);
    if goRowSelect in Options then
    begin
      oldsel.Left := self.FixedCols;
      oldsel.Right := self.ColCount - 1;
    end;
    if (sel.Left <> oldsel.Left) or (sel.right <> oldsel.right) or
       (sel.top <> oldsel.top)   or (sel.bottom <> oldsel.bottom) then
    begin
      t := mousecoord(x, y);
      t1 := mousecoord(FMouseXY.x, FMouseXY.y);
      if (t1.x < t.x) or (t1.y < t.y) then
      begin
        if (oldsel.Left < oldsel.Right) then
        begin
          sel.left := oldsel.Left;
          oldsel.Left := oldsel.Right;
          oldsel.Right := sel.Left;
        end;
        if (oldsel.top < oldsel.bottom) then
        begin
          sel.Left := oldsel.top;
          oldsel.top := oldsel.bottom;
          oldsel.bottom := sel.Left;
        end;
      end;
      Selection := oldsel;
      invalidate;
    end;
  end;
  RowSelectYY(VK_DOWN);
end;

procedure TZColorStringGrid.HideEditor;
begin
  ZInplaceEditor.Hide;
end;

procedure TZColorStringGrid.ShowEditorFirst;
begin
  ShowEditor;
  ZInplaceEditor.text := inherited GetEditText(Col, Row);
end;

//показать эдитор
procedure TZColorStringGrid.ShowEditor;
var
  t: integer;

  // изменить размеры эдитора
  //   _w - width
  //   _h - height
  procedure SetSizeEditor(_w, _h: integer);
  var
    t: integer;

  begin
    //GetSystemMetrics - получает системные настройки
    //        SM_CXHSCROLL - высота горизонтальной прокрутки
    //        SM_CXVSCROLL - ширина вертикальной прокрутки
    // отнимать 2 из-за линий
    // Если полосы прокрутки не видны - то их не учитываем

    if (GetWindowlong(self.Handle, GWL_STYLE) and WS_VSCROLL) <> 0 then
      t := GetSystemMetrics(SM_CXVSCROLL)
    else
      t := 0;

    if _w + ZInplaceEditor.Left > left + width - t - 2 then
      _w :=width + left - ZInplaceEditor.left - t - 2;
    ZInplaceEditor.width := _w;

    if (GetWindowlong(self.Handle, GWL_STYLE) and WS_HSCROLL) <> 0 then
      t := GetSystemMetrics(SM_CXHSCROLL)
    else
      t := 0;

    if _h + ZInplaceEditor.Top > top + height - t then
      _h :=height + top - ZInplaceEditor.Top - t - 2;
    ZInplaceEditor.Height := _h;
  end;

begin
  if csDesigning in ComponentState then exit;
  if (goEditing in Options)and(not (goRowSelect in Options)) then
  begin
    t := MergeCells.InMergeRange(Col, row);
    if ZInplaceEditor.parent <> Parent then
      ZInplaceEditor.parent := Parent;
    if t >= 0 then
    begin
      col := MergeCells.Items[t].left;
      row := MergeCells.Items[t].top;
    end;
    if (CellRect(Col, Row).left = 0) and (CellRect(Col, Row).Right = 0) then
    begin
      HideEditor;
      exit;
    end;
    ZInplaceEditor.Lines.Text := Cells[Col, Row];
    ZInplaceEditor.Left := CellRect(Col, Row).left  + self.left + 2;
    ZInplaceEditor.top := CellRect(Col, Row).Top + self.top + 2;

    if ZInplaceEditor.BorderStyle <> InplaceEditorOptions.BorderStyle then
      ZInplaceEditor.BorderStyle := InplaceEditorOptions.BorderStyle;

    ZInplaceEditor.Font := CellStyle[Col, Row].Font;
    if InplaceEditorOptions.UseCellStyle then
    begin
      ZInplaceEditor.Color := CellStyle[Col, Row].BGColor;
      ZInplaceEditor.Alignment := CellStyle[Col, Row].HorizontalAlignment;
    end else
    begin
      ZInplaceEditor.Font.color := InplaceEditorOptions.FontColor;
      ZInplaceEditor.Color := InplaceEditorOptions.BGColor;
      ZInplaceEditor.Alignment := InplaceEditorOptions.Alignment;
      ZInplaceEditor.WordWrap := InplaceEditorOptions.WordWrap;
    end;
    if t>=0 then
    begin
      if (ZInplaceEditor.Height + ZInplaceEditor.Top > self.Top + self.Height) or
         (ZInplaceEditor.Width + ZInplaceEditor.left > self.Left + self.Width) then
        ZInplaceEditor.Visible := false
      else
        SetSizeEditor(MergeCells.GetWidthArea(t),MergeCells.GetHeightArea(t));
      if ZInplaceEditor.Height + ZInplaceEditor.Top > self.Top + self.Height then
      begin
        //так надо! (иначе проблемы при TopLeftChanged)
        ZInplaceEditor.Height := 0;
        if row < rowcount-1 then row := row + 1;
      end;
      if (ZInplaceEditor.Width + ZInplaceEditor.left > self.Left + self.Width) then
      begin
        //так надо! (иначе проблемы при TopLeftChanged)
        ZInplaceEditor.Width := 0;
        col := MergeCells.Items[t].Right;
      end;
    end
    else
      SetSizeEditor(self.ColWidths[Col], self.RowHeights[Row]);
    ZInplaceEditor.show;
    ZInplaceEditor.SetFocus;
  end;
end;

function TZColorStringGrid.SelectCell(ACol, ARow: Longint): Boolean;
var
  t: integer;

begin
  //Если покинули(пришли в) объединённую ячейку - перерисовать
  t := MergeCells.InMergeRange(Col, Row);
  if t>=0 then Invalidate else
  begin
    t := MergeCells.InMergeRange(ACol, ARow);
    if (t>=0) then Invalidate;
  end;
  if (goEditing in Options)and(not (goRowSelect in Options))and (goAlwaysShowEditor in Options) then
    HideEditor
  else
  begin
    if ZInplaceEditor.Visible then
    begin
      HideEditor;
      if CanFocus then SetFocus;
    end;
  end;
  result := inherited SelectCell(ACol, ARow);
end;

destructor TZColorStringGrid.Destroy();
var
  i,j: integer;

begin
  FSelectedColors.Free;
  FLineDesign.Free;
  for i :=0 to ColCount - 1 do
  begin
    for j := 0 to self.RowCount - 1 do
      FCellStyle[i][j].free;
    setlength(FCellStyle[i],0);
  end;
  setlength(FCellStyle, 0);
  FDefaultCellParam.Free;
  FDefaultFixedCellParam.Free;
  FMergeCells.Free;
  FBmp.free;
  FInplaceEditorOptions.Free;
  inherited Destroy();
end;

procedure Register;
begin
  RegisterComponents('ZColor', [TZColorStringGrid]);
end;

end.

//В конце понял, что лучше бы делал с нуля, а не
//наследовал от TStringGrid. Но переделывать лень.

//Copyright © 2008 Неборак Руслан Владимирович
