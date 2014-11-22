//****************************************************************
// ZEXMLSS  (Z Excel XML SpreadSheet) - невизуальный компонент,
// предназначенный для сохранения данных в формате excel xml.
// Накалякано в Мозыре в 2009 году
// Автор:  Неборак Руслан Владимирович (Ruslan V. Neborak)
// e-mail: avemey(мяу)tut(точка)by
// URL:    http://avemey.com
// Ver:    0.0.7
// Лицензия: zlib
// Last update: 2014.07.20
//----------------------------------------------------------------
// This software is provided "as-is", without any express or implied warranty.
// In no event will the authors be held liable for any damages arising from the
// use of this software.
//****************************************************************
unit zexmlss;

{$I zexml.inc}
{$I compver.inc}

{$IFDEF FPC}
  {$mode objfpc}{$H+}
{$ENDIF}

interface

// т.к. в Delphi 5 нету таких полезных функций как AnsiToUtf8 и Utf8ToAnsi
// подключаем sysd7.
// VER130 - Borland Delphi 5.0
uses
  classes, Sysutils,
  graphics,
  {$IFNDEF FPC}
  windows,
  {$ELSE}
  LCLType,
  LResources,
  {$ENDIF}
  zsspxml
  ;

const
   _PointToMM:    real = 0.3528;  // 1 типографский пункт = 0.3528 мм

type
  //тип данных ячейки
  TZCellType = (ZENumber, ZEDateTime, ZEBoolean, ZEString, ZEError);
      const ZEAnsiString = ZEString deprecated {$IFDEF USE_DEPRECATED_STRING}'use ZEString'{$ENDIF}; // backward compatibility
type
  //Стиль начертания линий рамки ячейки
  TZBorderType = (ZENone, ZEContinuous, ZEDot, ZEDash, ZEDashDot, ZEDashDotDot,ZESlantDashDot, ZEDouble);

  //Горизонтальное выравнивание
  TZHorizontalAlignment = (ZHAutomatic, ZHLeft, ZHCenter, ZHRight, ZHFill, ZHJustify, ZHCenterAcrossSelection, ZHDistributed, ZHJustifyDistributed);

  //вертикальное выравнивание в ячейке
  TZVerticalAlignment = (ZVAutomatic, ZVTop, ZVBottom, ZVCenter, ZVJustify, ZVDistributed, ZVJustifyDistributed);

  //Шаблон заливки ячейки
  TZCellPattern = (ZPNone, ZPSolid, ZPGray75, ZPGray50, ZPGray25, ZPGray125, ZPGray0625, ZPHorzStripe, ZPVertStripe,
                  ZPReverseDiagStripe, ZPDiagStripe, ZPDiagCross, ZPThickDiagCross, ZPThinHorzStripe, ZPThinVertStripe,
                  ZPThinReverseDiagStripe, ZPThinDiagStripe, ZPThinHorzCross, ZPThinDiagCross);

  //Закрепление строк/столбцов
  TZSplitMode = (ZSplitNone, ZSplitFrozen, ZSplitSplit);

  //ячейка
  TZCell = class(TPersistent)
  private
    FFormula: string;
    FData: string;
    FHref: string;
    FHRefScreenTip: string;
    FComment: string;
    FCommentAuthor: string;
    FAlwaysShowComment: boolean; //default = false;
    FShowComment: boolean;       //default = false
    FCellType: TZCellType;
    FCellStyle: integer;
    function GetDataAsDouble: double;
    procedure SetDataAsDouble(const Value: double);
    procedure SetDataAsInteger(const Value: integer);
    function GetDataAsInteger: integer;
  public
    constructor Create();virtual;
    procedure Assign(Source: TPersistent); override;
    procedure Clear();
    property AlwaysShowComment: boolean read FAlwaysShowComment write FAlwaysShowComment default false;
    property Comment: string read FComment write FComment;      //примечание
    property CommentAuthor: string read FCommentAuthor write FCommentAuthor;// автор примечания
    property CellStyle: integer read FCellStyle write FCellStyle default -1;
    property CellType: TZCellType read FCellType write FCellType default ZEString; //тип данных ячейки
    property Data: string read FData write FData;              //отображаемое содержимое ячейки
    property Formula: string read FFormula write FFormula;     //формула в стиле R1C1
    property HRef: string read FHref write FHref;              //гиперссылка
    property HRefScreenTip: string read FHRefScreenTip write FHRefScreenTip; //подпись гиперссылки
    property ShowComment: boolean read FShowComment write FShowComment default false;

    property AsDouble: double read GetDataAsDouble write SetDataAsDouble;
    property AsInteger: integer read GetDataAsInteger write SetDataAsInteger;
  end;

  //стиль линии границы
  TZBorderStyle = class (TPersistent)
  private
    FLineStyle: TZBorderType;
    FWeight: byte;
    FColor: TColor;
    procedure SetLineStyle(const Value: TZBorderType);
    procedure SetWeight(const Value: byte);
    procedure SetColor(const Value: TColor);
  public
    constructor Create();virtual;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    property LineStyle: TZBorderType read FLineStyle write SetLineStyle default ZENone;
    property Weight: byte read FWeight write SetWeight default 0;
    property Color: TColor read FColor write SetColor default ClBlack;
  end;

  //Рамка ячейки
  //   0 - left           левая граница
  //   1 - Top            верхняя граница
  //   2 - Right          правая граница
  //   3 - Bottom         нижняя граница
  //   4 - DiagonalLeft   диагоняль от верхнего левого угла до нижнего правого
  //   5 - DiagonalRight  диагоняль от нижнего левого угла до правого верхнего
  TZBorder = class (TPersistent)
  private
    FBorder: array [0..5] of TZBorderStyle;
    procedure SetBorder(Num: integer; Const Value: TZBorderStyle);
    function GetBorder(Num: integer):TZBorderStyle;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent);override;
    property Border[Num: integer]: TZBorderStyle read GetBorder write SetBorder; default;
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    property Left         : TZBorderStyle index 0 read GetBorder write SetBorder;
    property Top          : TZBorderStyle index 1 read GetBorder write SetBorder;
    property Right        : TZBorderStyle index 2 read GetBorder write SetBorder;
    property Bottom       : TZBorderStyle index 3 read GetBorder write SetBorder;
    property DiagonalLeft : TZBorderStyle index 4 read GetBorder write SetBorder;
    property DiagonalRight: TZBorderStyle index 5 read GetBorder write SetBorder;
  end;

  /// Угол поворота текста в ячейке. Целое со знаком.
  ///    Базовый диапазон -90 .. +90, расширенный -180 .. +180 (градусов)
  ///    Значения из расширенного диапазона НЕ МОГУТ работать в Microsoft Excel!
  /// Определение по XML SS совпадает с базовым
  /// Определение по OpenDocument 1.2 позволяет любое дробное число
  ///    в разных еденицах измерения (градусы, грады, радианы).
  ///    практически используется целое от 0 до 359
  /// Определение по OfficeXML наиболее запутанное:
  ///    254 и 255 - специальные значения,
  ///        255 работает в Excel хотя в стандарте не описано
  ///        254 принимается, но после открытия файла не работает
  ///    от 0 до +90 соответствует обоим вышеописанным стандартам
  ///    от +91 до +180 соответствует значениям -1 .. -90 [ = 90 - value]
  ///
  /// 1. MSDN aa140066.aspx: XML Spreadsheet Reference => ss:Rotate
  /// 2. OASIS OpenDocument 1.2 (Sep.2011) Part 1 => 18.3 Other Datatypes => 18.3.1 angle
  /// 3. ECMA-376 4th Edition (Dec.2012) Part 1 => 18.8 Styles => 18.8.1 alignment
  ///
  ///  An angle of the rotation (direction) for a text within a cell. Signed Integer;
  ///    Nominative range is -90 .. +90, extended is -180 .. +180 (in degrees)
  ///       Extended range values can not be loaded by Microsoft Excel!
  /// XML SS specifications matches the nominative range.
  /// OpenDocument 1.2 specification, atchign mathematics and SVG, permits FP
  ///    numbers in different unit (deg, grad, rad) with degrees by default;
  ///    practical implementations use non-negative integer numbers in 0..359 range.
  /// OfficeXML specification is the most perplexed
  ///    254 and 255 are special values, and:
  ///        255 works within MS Excel despite being not defined in the specs;
  ///        254 can be opened by excel, but such a cell can not be rendered.
  ///    The 0 to +90 range matches both standards aforementioned
  ///    The +91 to +180 range matches -1 to -90 range. [ = 90 - value]
  ///
  TZCellTextRotate = -180 .. +359;

  //выравнивание
        //ReadingOrder - имхо, не нужно ^_^
  TZAlignment = class (TPersistent)
  private
    FHorizontal: TZHorizontalAlignment; //горизонтальное выравнивание
    FIndent: integer;                   //отступ
    FRotate: TZCellTextRotate;          //угол поворота
    FShrinkToFit: boolean;              //true - уменьшает размер шрифта, чтобы текст
                                        //   поместился в ячейку,
                                        // false - текст не уменьшается
    FVertical: TZVerticalAlignment;     //Вертикальное выравнивание
    FVerticalText: boolean;             //true  - текст по одной букве в строке вертикально
                                        //false - по дефолту
    FWrapText: boolean;                 //перенос текста
    procedure SetHorizontal(const Value: TZHorizontalAlignment);
    procedure SetIndent(const Value: integer);
    procedure SetRotate(const Value: TZCellTextRotate);
    procedure SetShrinkToFit(const Value: boolean);
    procedure SetVertical(const Value: TZVerticalAlignment);
    procedure SetVerticalText(const Value:boolean);
    procedure SetWrapText(const Value: boolean);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    property Horizontal: TZHorizontalAlignment read FHorizontal write SetHorizontal default ZHAutomatic;
    property Indent: integer read FIndent write SetIndent default 0;
    property Rotate: TZCellTextRotate read FRotate write SetRotate;
    property ShrinkToFit: boolean read FShrinkToFit write SetShrinkToFit default false;
    property Vertical: TZVerticalAlignment read FVertical write SetVertical default ZVAutomatic;
    property VerticalText: boolean read FVerticalText write SetVerticalText default false;
    property WrapText: boolean read FWrapText write SetWrapText default false;
  end;

  //стиль ячейки
  TZStyle = class (TPersistent)
  private
    FBorder: TZBorder;
    FAlignment: TZAlignment;
    FFont: TFont;   //шрифт ячейки
    FBGColor: TColor;
    FPatternColor: TColor;
    FCellPattern: TZCellPattern;
    FNumberFormat: string;
    FProtect: boolean;
    FHideFormula: boolean;
    procedure SetFont(const Value: TFont);
    procedure SetBorder(const Value: TZBorder);
    procedure SetAlignment(const Value: TZAlignment);
    procedure SetBGColor(const Value: TColor);
    procedure SetPatternColor(const Value: TColor);
    procedure SetCellPattern(const Value: TZCellPattern);
  protected
    procedure SetNumberFormat(const Value: string); virtual;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    property Font: TFont read FFont write SetFont;
    property Border:TZBorder read FBorder write SetBorder;
    property Alignment: TZAlignment read FAlignment write SetAlignment;
    property BGColor: TColor read FBGColor write SetBGColor default clWindow;
    property PatternColor: TColor read FPatternColor write SetPatternColor default clWindow;
    property Protect: boolean read FProtect write FProtect default true;
    property HideFormula: boolean read FHideFormula write FHideFormula default false;
    property CellPattern: TZCellPattern read FCellPattern write SetCellPattern default ZPNone;
    property NumberFormat: string read FNumberFormat write SetNumberFormat;
  end;

  //стили
  TZStyles = class (TPersistent)
  private
    FDefaultStyle: TZStyle;
    FStyles: array of TZStyle;
    FCount: integer;
    procedure SetDefaultStyle(const Value: TZStyle);
    function  GetStyle(num: integer): TZStyle;
    procedure SetStyle(num: integer; const Value: TZStyle);
    procedure SetCount(const Value: integer);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(const Style: TZStyle; CheckMatch: boolean = false): integer;
    procedure Assign(Source: TPersistent); override;
    procedure Clear(); virtual;
    function DeleteStyle(num: integer):integer; virtual;
    function Find(const Style: TZStyle): integer;
    property Items[num: integer]: TZStyle read GetStyle write SetStyle; default;
    property Count: integer read FCount write SetCount;
  published
    property DefaultStyle: TZStyle read FDefaultStyle write SetDefaultStyle;
  end;

  TZSheet = class;

  //Объединённые области
  TZMergeCells = class
  private
    FSheet: TZSheet;
    FCount: integer;            //кол-во объединённых областей
    FMergeArea: Array of TRect;
    function GetItem(Num: integer): TRect;
  public
    constructor Create(ASheet: TZSheet);virtual;
    destructor Destroy();override;
    function AddRect(Rct:TRect): byte;
    function AddRectXY(x1,y1,x2,y2: integer): byte;
    function DeleteItem(num: integer): boolean;
    function InLeftTopCorner(ACol, ARow: integer): integer;
    function InMergeRange(ACol, ARow: integer): integer;
    procedure Clear();
    property Count: integer read FCount;
    property Items[Num: Integer]: TRect read GetItem; default;
  end;

  TZCellColumn = array of TZCell;

  TZEXMLSS = class;

  //общие настройки Row / Column
  TZRowColOptions = class (TPersistent)
  private
    FSheet: TZSheet;
    FHidden: boolean;
    FStyleID: integer;
    FSize: real;
    FAuto: boolean;
    FBreaked: boolean;
  protected
    function  GetAuto(): boolean;
    procedure SetAuto(Value: boolean);
    function  GetSizePoint(): real;
    procedure SetSizePoint(Value: real);
    function  GetSizeMM(): real;
    procedure SetSizeMM(Value: real);
    function  GetSizePix(): integer; virtual; abstract;
    procedure SetSizePix(Value: integer); virtual; abstract;
  public
    constructor Create(ASheet: TZSheet); virtual;
    procedure Assign(Source: TPersistent); override;
    property Hidden: boolean read FHidden write FHidden default false;
    property StyleID: integer read FStyleID write FStyleID default -1;
    property Breaked: boolean read FBreaked write FBreaked default false;
  end;

  //Columns
  TZColOptions = class (TZRowColOptions)
  protected
    function  GetSizePix(): integer; override;
    procedure SetSizePix(Value: integer); override;
  public
    constructor Create(ASheet: TZSheet); override;
    property AutoFitWidth: boolean read GetAuto write SetAuto;
    property Width: real read GetSizePoint write SetSizePoint;
    property WidthMM: real read GetSizeMM write SetSizeMM;
    property WidthPix: integer read GetSizePix write SetSizePix;
  end;

  //Rows
  TZRowOptions = class (TZRowColOptions)
  protected
    function  GetSizePix(): integer; override;
    procedure SetSizePix(Value: integer); override;
  public
    constructor Create(ASheet: TZSheet); override;
    property AutoFitHeight: boolean read GetAuto write SetAuto;
    property Height: real read GetSizePoint write SetSizePoint;
    property HeightMM: real read GetSizeMM write SetSizeMM;
    property HeightPix: integer read GetSizePix write SetSizePix;
  end;

  /// repeating columns or rows when printing the sheet
  ///   implemented as binary mess in XLSX
  ///   implemented as named range and overriding cell names in XML SS
  ///   implemented as special container tags in OpenDocument
  TZSheetPrintTitles = class(TPersistent)
  private
    FOwner: TZSheet;
    FColumns: boolean;

    FActive: boolean;
    FTill: word;
    FFrom: word;
    procedure SetActive(const Value: boolean);
    procedure SetFrom(const Value: word);
    procedure SetTill(const Value: word);
    function  Valid(const AFrom, ATill: word): boolean;
    procedure RequireValid(const AFrom, ATill: word);
  public
    procedure Assign(Source: TPersistent); override;
    constructor Create(const owner: TZSheet; const ForColumns: boolean);

    function ToString: string; {$IfDef Delphi_Unicode} override; {$EndIf} {$IFDEF Z_FPC_USE_TOSTRING} override; {$ENDIF}
// introduced in D2009 according to http://blog.marcocantu.com/blog/6hidden_delphi2009.html
// introduced don't know when in FPC
  published

    property From: word read FFrom write SetFrom;
    property Till: word read FTill write SetTill;
    property Active: boolean read FActive write SetActive;
  end;

  //Footer/Header
  TZSheetFooterHeader = class (TPersistent)
  private
    FDataLeft: string;
    FData: string;
    FDataRight: string;
  public
    constructor Create();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    property DataLeft: string read FDataLeft write FDataLeft;
    property Data: string read FData write FData;
    property DataRight: string read FDataRight write FDataRight;
  end;

  //WorkSheetOptions
  TZSheetOptions = class (TPersistent)
  private
    // достали дюймы! Будем измерять в ММ!
    // кстати, может, вместо word использовать Cardinal?
    FActiveCol: word;
    FActiveRow: word;
    FMarginBottom: word;
    FMarginLeft: word;
    FMarginTop: word;
    FMarginRight: word;
    FPortraitOrientation: boolean;
    FCenterHorizontal: boolean;
    FCenterVertical: boolean;
    FStartPageNumber: integer;
    FHeaderMargin: word;
    FFooterMargin: word;
    FIsEvenFooterEqual: boolean;      //Is the footer on even and odd pages the same?
    FIsEvenHeaderEqual: boolean;      //Is the header on even and odd pages the same?

    FHeader: TZSheetFooterHeader;     //Header for all pages. When IsEvenFooterEqual = false - only for odd pages.
    FFooter: TZSheetFooterHeader;     //Footer for all pages. When IsEvenFooterEqual = false - only for odd pages.
    FEvenHeader: TZSheetFooterHeader; //Header for even pages. Used only if IsEvenFooterEqual = false
    FEvenFooter: TZSheetFooterHeader; //Footer for even pages. Used only if IsEvenFooterEqual = false

    FHeaderBGColor: TColor;
    FFooterBGColor: TColor;

    FPaperSize: byte;
    FPaperWidth: integer;
    FPaperHeight: integer;
    FSplitVerticalMode: TZSplitMode;
    FSplitHorizontalMode: TZSplitMode;
    FSplitVerticalValue: integer;       //Вроде можно вводить отрицательные
    FSplitHorizontalValue: integer;     //Измеряться будут:
                                        //    в пикселях, если SplitMode = ZSplitSplit
                                        //    в кол-ве строк/столбцов, если SplitMode = ZSplitFrozen
                                        // Если SplitMode = ZSplitNone, то фиксация столбцов/ячеек не работает
    function GetHeaderData(): string;
    procedure SetHeaderData(Value: string);
    function GetFooterData(): string;
    procedure SetfooterData(Value: string);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
  published
    property ActiveCol: word read FActiveCol write FActiveCol default 0;
    property ActiveRow: word read FActiveRow write FActiveRow default 0;
    property MarginBottom: word read FMarginBottom write FMarginBottom default 25;
    property MarginLeft: word read FMarginLeft write FMarginLeft default 20;
    property MarginTop: word read FMarginTop write FMarginTop default 25;
    property MarginRight: word read FMarginRight write FMarginRight default 20;
    property PaperSize: byte read FPaperSize write FPaperSize default 9; // A4
    property PaperWidth: integer read FPaperWidth write FPaperWidth default 0;    //User defined paper width in mm.
                                                                                  // used only for PaperSize = 0!
    property PaperHeight: integer read FPaperHeight write FPaperHeight default 0; //User defined paper height in mm.
                                                                                  // used only for PaperSize = 0!
    property PortraitOrientation: boolean read FPortraitOrientation write FPortraitOrientation default true;
    property CenterHorizontal: boolean read FCenterHorizontal write FCenterHorizontal default false;
    property CenterVertical: boolean read FCenterVertical write FCenterVertical default false;
    property StartPageNumber: integer read FStartPageNumber write FStartPageNumber default 1;
    property HeaderMargin: word read FHeaderMargin write FHeaderMargin default 13;
    property FooterMargin: word read FFooterMargin write FFooterMargin default 13;
    property IsEvenFooterEqual: boolean read FIsEvenFooterEqual write FIsEvenFooterEqual default true;
    property IsEvenHeaderEqual: boolean read FIsEvenHeaderEqual write FIsEvenHeaderEqual default true;
    property HeaderData: string read GetHeaderData write SetHeaderData;
    property FooterData: string read GetFooterData write SetFooterData;
    property Header: TZSheetFooterHeader read FHeader;
    property Footer: TZSheetFooterHeader read FFooter;
    property EvenHeader: TZSheetFooterHeader read FEvenHeader;
    property EvenFooter: TZSheetFooterHeader read FEvenFooter;
    property HeaderBGColor: TColor read FHeaderBGColor write FHeaderBGColor default clWindow;
    property FooterBGColor: TColor read FFooterBGColor write FFooterBGColor default clWindow;
    property SplitVerticalMode: TZSplitMode read FSplitVerticalMode write FSplitVerticalMode default ZSplitNone;
    property SplitHorizontalMode: TZSplitMode read FSplitHorizontalMode write FSplitHorizontalMode default ZSplitNone;
    property SplitVerticalValue: integer read FSplitVerticalValue write FSplitVerticalValue;
    property SplitHorizontalValue: integer read FSplitHorizontalValue write FSplitHorizontalValue;
  end;

  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}

  //Условие
  TZCondition = (ZCFIsTrueFormula,
                 ZCFCellContentIsBetween,
                 ZCFCellContentIsNotBetween,
                 ZCFCellContentOperator,
                 ZCFNumberValue,
                 ZCFString,
                 ZCFBoolTrue,
                 ZCFBoolFalse,
                 ZCFFormula,
                 ZCFContainsText,
                 ZCFNotContainsText,
                 ZCFBeginsWithText,
                 ZCFEndsWithText,
                 ZCFCellIsEmpty,
                 ZCFDuplicate,
                 ZCFUnique,
                 ZCFAboveAverage,
                 ZCFBellowAverrage,
                 ZCFAboveEqualAverage,
                 ZCFBelowEqualAverage,
                 ZCFTopElements,
                 ZCFBottomElements,
                 ZCFTopPercent,
                 ZCFBottomPercent,
                 ZCFIsError,
                 ZCFIsNoError
                );

  //Оператор для условного форматирования
  TZConditionalOperator = (ZCFOpGT,       //  > (Greater Than)
                           ZCFOpLT,       //  < (Less Than)
                           ZCFOpGTE,      //  >= (Greater or Equal)
                           ZCFOpLTE,      //  <= (Less or Equal)
                           ZCFOpEqual,    //  = (Equal)
                           ZCFOpNotEqual  //  <> (Not Equal)
                          );

  //Условное форматирование: стиль на условие (conditional formatting)
  TZConditionalStyleItem = class (TPersistent)
  private
    FCondition: TZCondition;                    //условие
    FConditionOperator: TZConditionalOperator;  //Оператор
    FValue1: String;
    FValue2: String;
    FApplyStyleID: integer;                     //номер применяемого стиля

                                                //Базовая ячейка (только для формул):
    FBaseCellPageIndex: integer;                //  Номер страницы для адреса базовой ячейки
                                                //    -1 - текущая страница
    FBaseCellRowIndex: integer;                 //  Номер строки
    FBaseCellColumnIndex: integer;              //  Номер столбца
  protected
  public
    constructor Create(); virtual;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property ApplyStyleID: integer read FApplyStyleID write FApplyStyleID;
    property BaseCellColumnIndex: integer read FBaseCellColumnIndex write FBaseCellColumnIndex;
    property BaseCellPageIndex: integer read FBaseCellPageIndex write FBaseCellPageIndex;
    property BaseCellRowIndex: integer read FBaseCellRowIndex write FBaseCellRowIndex;
    property Condition: TZCondition read FCondition write FCondition;
    property ConditionOperator: TZConditionalOperator read FConditionOperator write FConditionOperator;
    property Value1: String read FValue1 write FValue1;
    property Value2: String read FValue2 write FValue2;
  end;

  //Область для применения условного форматирования
  TZConditionalAreaItem = class (TPersistent)
  private
    FRow: integer;
    FColumn: integer;
    FWidth: integer;
    FHeight: integer;
    procedure SetRow(Value: integer);
    procedure SetColumn(Value: integer);
    procedure SetWidth(Value: integer);
    procedure SetHeight(Value: integer);
  protected
  public
    constructor Create(); overload; virtual;
    constructor Create(ColumnNum, RowNum, AreaWidth, AreaHeight: integer); overload; virtual;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Row: integer read FRow write SetRow;
    property Column: integer read FColumn write SetColumn;
    property Width: integer read FWidth write SetWidth;
    property Height: integer read FHeight write SetHeight;
  end;

  //Области для применения условного форматирования
  TZConditionalAreas = class (TPersistent)
  private
    FCount: integer;
    FItems: array of TZConditionalAreaItem;
    procedure SetCount(Value: integer);
    function GetItem(num: integer): TZConditionalAreaItem;
    procedure SetItem(num: integer; Value: TZConditionalAreaItem);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Add(); overload;
    procedure Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer); overload;
    procedure Assign(Source: TPersistent); override;
    procedure Delete(num: integer);
    function IsCellInArea(ColumnNum, RowNum: integer): boolean;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Count: integer read FCount write SetCount;
    property Items[num: integer]: TZConditionalAreaItem read GetItem write SetItem; default;
  end;

  //Условное форматирование: список условий
  TZConditionalStyle = class (TPersistent)
  private
    FCount: integer;                    //кол-во условий
    FMaxCount: integer;
    FAreas: TZConditionalAreas;
    FConditions: array of TZConditionalStyleItem;
    function GetItem(num: integer): TZConditionalStyleItem;
    procedure SetItem(num: integer; Value: TZConditionalStyleItem);
    procedure SetCount(value: integer);
    procedure SetAreas(Value: TZConditionalAreas);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalStyleItem; overload;
    function Add(StyleItem: TZConditionalStyleItem): TZConditionalStyleItem; overload;
    procedure Delete(num: integer);
    procedure Insert(num: integer); overload;
    procedure Insert(num: integer; StyleItem: TZConditionalStyleItem); overload;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Areas: TZConditionalAreas read FAreas write SetAreas;
    property Count: integer read FCount write SetCount;
    property Items[num: integer]: TZConditionalStyleItem read GetItem write SetItem; default;
  end;

  TZConditionalFormatting = class (TPersistent)
  private
    FStyles: array of TZConditionalStyle;
    FCount: integer;
    procedure SetCount(Value: integer);
    function GetItem(num: integer): TZConditionalStyle;
    procedure SetItem(num: integer; Value: TZConditionalStyle);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalStyle; overload;
    function Add(Style: TZConditionalStyle): TZConditionalStyle; overload;
    function Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer): TZConditionalStyle; overload;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Count: integer read FCount write SetCount;
    property Items[num: integer]: TZConditionalStyle read GetItem write SetItem; default;
  end;

  {$ENDIF} //ZUSE_CONDITIONAL_FORMATTING

  //лист документа
  TZSheet = class (TPersistent)
  private
    FStore: TZEXMLSS;
    FCells: array of TZCellColumn;
    FRows: array of TZRowOptions;
    FColumns: array of TZColOptions;
    FTitle: string;                     //заголовок листа
    FRowCount: integer;
    FColCount: integer;
    FTabColor: TColor;                  //цвет закладки
    FDefaultRowHeight: real;
    FDefaultColWidth: real;
    FMergeCells: TZMergeCells;
    FProtect: boolean;
    FRightToLeft: boolean;
    FSheetOptions: TZSheetOptions;
    FSelected: boolean;
    FPrintRows, FPrintCols: TZSheetPrintTitles;

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    FConditionalFormatting: TZConditionalFormatting;
    procedure SetConditionalFormatting(Value: TZConditionalFormatting);
    {$ENDIF}

    procedure SetColumn(num: integer; const Value:TZColOptions);
    function  GetColumn(num: integer): TZColOptions;
    procedure SetRow(num: integer; const Value:TZRowOptions);
    function  GetRow(num: integer): TZRowOptions;
    function  GetSheetOptions(): TZSheetOptions;
    procedure SetSheetOptions(Value: TZSheetOptions);
    procedure SetPrintCols(const Value: TZSheetPrintTitles);
    procedure SetPrintRows(const Value: TZSheetPrintTitles);
  protected
    procedure SetColWidth(num: integer; const Value: real); virtual;
    function  GetColWidth(num: integer): real; virtual;
    procedure SetRowHeight(num: integer; const Value: real); virtual;
    function  GetRowHeight(num: integer): real; virtual;
    procedure SetDefaultColWidth(const Value: real); virtual;
    procedure SetDefaultRowHeight(const Value: real); virtual;
    function  GetCell(ACol, ARow: integer): TZCell; virtual;
    procedure SetCell(ACol, ARow: integer; const Value: TZCell); virtual;
    function  GetRowCount: integer; virtual;
    procedure SetRowCount(const Value: integer); virtual;
    function  GetColCount: integer; virtual;
    procedure SetColCount(const Value: integer); virtual;
  public
    constructor Create(AStore: TZEXMLSS); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    procedure Clear(); virtual;
    property ColWidths[num: integer]: real read GetColWidth write SetColWidth;
    property Columns[num: integer]: TZColOptions read GetColumn write SetColumn;
    property Rows[num: integer]: TZRowOptions read GetRow write SetRow;
    property RowHeights[num: integer]: real read GetRowHeight write SetRowHeight;
    property DefaultColWidth: real read FDefaultColwidth write SetDefaultColWidth;// default 48;
    property DefaultRowHeight: real read FDefaultRowHeight write SetDefaultRowHeight;// default 12.75;
    property Cell[ACol, ARow: integer]: TZCell read GetCell write SetCell; default;
    property Protect: boolean read FProtect write FProtect default false; //защищён ли лист от изменения
    property TabColor: TColor read FTabColor write FTabColor default ClWindow;
    property Title: string read FTitle write FTitle;
    property RowCount: integer read GetRowCount write SetRowCount;
    property RightToLeft: boolean read FRightToLeft write FRightToLeft default false;
    property ColCount: integer read GetColCount write SetColCount;
    property MergeCells: TZMergeCells read FMergeCells write FMergeCells;
    property SheetOptions: TZSheetOptions read GetSheetOptions write SetSheetOptions;
    property Selected: boolean read FSelected write FSelected;

    property RowsToRepeat: TZSheetPrintTitles read FPrintRows write SetPrintRows;
    property ColsToRepeat: TZSheetPrintTitles read FPrintCols write SetPrintCols;

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    property ConditionalFormatting: TZConditionalFormatting read FConditionalFormatting write SetConditionalFormatting;
    {$ENDIF}
  end;

  //Страницы
  TZSheets = class (TPersistent)
  private
    FStore: TZEXMLSS;
    FSheets: array of TZSheet;
    FCount : integer;
    procedure SetSheetCount(const Value: integer);
    procedure SetSheet(num: integer; Const Value: TZSheet);
    function  GetSheet(num: integer): TZSheet;
  protected
  public
    constructor Create(AStore: TZEXMLSS); virtual;
    destructor  Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    property Count: integer read FCount write SetSheetCount;
    property Sheet[num: integer]: TZSheet read GetSheet write SetSheet; default;
  end;

  //Свойства документа
  TZEXMLDocumentProperties = class(TPersistent)
  private
    FAuthor      : string;
    FLastAuthor  : string;
    FCreated     : TdateTime;
    FCompany     : string;
    FVersion     : string; // - should be integer by Spec but hardcoded float in real MS Office apps
    FWindowHeight: word;
    FWindowWidth : word;
    FWindowTopX  : integer;
    FWindowTopY  : integer;
    FModeR1C1    : boolean;
  protected
    procedure SetAuthor(const Value: string);
    procedure SetLastAuthor(const Value: string);
    procedure SetCompany(const Value: string);
    procedure SetVersion(const Value: string);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
  published
    property Author: string read FAuthor write SetAuthor;
    property LastAuthor: string read FLastAuthor write SetLastAuthor;
    property Created: TdateTime read FCreated write FCreated;
    property Company: string read FCompany write SetCompany;
    property Version: string read FVersion write SetVersion;
    property ModeR1C1: boolean read FModeR1C1 write FModeR1C1 default false;
    property WindowHeight: word read FWindowHeight write FWindowHeight default 20000;
    property WindowWidth: word read FWindowWidth write FWindowWidth default 20000;
    property WindowTopX: integer read FWindowTopX write FWindowTopX default 150;
    property WindowTopY: integer read FWindowTopY write FWindowTopY default 150;
  end;

  TZEXMLSS = class (TComponent)
  private
    FSheets: TZSheets;
    FDocumentProperties: TZEXMLDocumentProperties;
    FStyles: TZStyles;
    FHorPixelSize: real;
    FVertPixelSize: real;
    FDefaultSheetOptions: TZSheetOptions;
    procedure SetHorPixelSize(Value: real);
    procedure SetVertPixelSize(Value: real);
    function  GetDefaultSheetOptions(): TZSheetOptions;
    procedure SetDefaultSheetOptions(Value: TZSheetOptions);
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy();override;
    procedure Assign(Source: TPersistent); override;
    procedure GetPixelSize(hdc: HWND);  // получает HorPixelSize и VertPixelSize
    //ИМХО, лучще сделать функции сохранения/загрузки в другом модуле
    //function SaveToFile(const FileName: ansistring; const SheetsNumbers: array of integer; const SheetsNames: array of ansistring; CodePage: byte {$IFDEF VER130}{$ELSE} = 0{$ENDIF}): integer; overload; virtual;
    //function SaveToStream(Stream: TStream; const SheetsNumbers: array of integer; const SheetsNames: array of ansistring; CodePage: byte {$IFDEF VER130}{$ELSE} = 0{$ENDIF}): integer; overload; virtual;
    //function SaveToFile(const FileName: ansistring; const SheetsNumbers: array of integer; CodePage: byte = 0): integer; overload; virtual;
    //function SaveToStream(Stream: TStream; const SheetsNumbers: array of integer; CodePage: byte = 0): integer; overload; virtual;
    //function SaveToFile(const FileName: ansistring; CodePage: byte = 0): integer; overload; virtual;
    //function SaveToStream(Stream: TStream; CodePage: byte = 0): integer; overload; virtual;
    property Sheets: TZSheets read FSheets write FSheets;
  published
    property Styles: TZStyles read FStyles write FStyles;
    property DefaultSheetOptions: TZSheetOptions read GetDefaultSheetOptions write SetDefaultSheetOptions;
    property DocumentProperties: TZEXMLDocumentProperties read FDocumentProperties write FDocumentProperties;
    property HorPixelSize: real read FHorPixelSize write SetHorPixelSize;  //размер пикселя по горизонтали
    property VertPixelSize: real read FVertPixelSize write SetVertPixelSize;  //размер пикселя по вертикали
  end;

procedure Register();

//Проверяет строку на правильность, если что-то не правильно - правит
procedure CorrectStrForXML(const St: string; var Corrected: string; var UseXMLNS: boolean);

//переводит TColor в Hex RGB
function ColorToHTMLHex(Color: TColor): string;

//переводит Hex RGB (string) в TColor
function HTMLHexToColor(value: string): TColor;

//Перевести пиксели в типографский пункт (point)
function PixelToPoint(inPixel: integer; PixelSizeMM: real = 0.265): real;

//Перевести типографский пункт (point) в пиксели
function PointToPixel(inPoint: real; PixelSizeMM: real = 0.265): integer;

//Перевести типографский пункт (point) в мм
function PointToMM(inPoint: real): real;

//Перевести мм в типографский пункт (point)
function MMToPoint(inMM: real): real;

//HorizontalAligment to Str
function HAlToStr(HA: TZHorizontalAlignment): string;

//Vertical Aligment to Str
function VAlToStr(VA: TZVerticalAlignment): string;

function ZBorderTypeToStr(ZB: TZBorderType): string;

function ZCellPatternToStr(pp: TZCellPattern): string;

function ZCellTypeToStr(pp: TZCellType): string;

//Str to HorizontalAligment
function StrToHal(Value: string): TZHorizontalAlignment;

//Str to Vertical Aligment
function StrToVAl(Value: string): TZVerticalAlignment;

function StrToZBorderType(Value: string): TZBorderType;

function StrToZCellPattern(Value: string): TZCellPattern;

function StrToZCellType(Value: string): TZCellType;

implementation

//Проверяет строку на правильность, если что-то не правильно - правит
//Input
//      St: string - исходная строка
//      var Corrected: string - исправленная строка
//      var UseXMLNS: boolean - есть ли в строке правильные парные тэги
//                              true - есть
procedure CorrectStrForXML(const St: string; var Corrected: string; var UseXMLNS: boolean);
type
  StackData = record
    TextBeforeTag: string;
    TagName: string;
    RawTag: string;
    isBad: boolean;
    isClose: boolean;
  end;

var
  i, kol, _length, t: integer;
  s: string;
  Stack: array of StackData;

  //проверка тэга
  procedure _CheckTag(var _num: integer; const _s: string; var D: StackData);
  var
    i, l, f, t: integer;
    s: string;
    _param: string;
    _Start: boolean;
    _zz: string;

  begin
    l := length(_s);
    f := 0;
    s := '';
    _Start := false;
    _param := '';
    _zz := '';
    for i := _num + 1 to length(_s) do
    if _s[i] = '>' then
    begin
      f := i;
      break;
    end;
    //убитый тэг
    if f = 0 then
    begin
      D.isBad := true;
      D.RawTag := copy(_s, _num, l - _num + 1);
      _num := l;
    end else
    begin
      for i := _num + 1 to f do
      begin
        case _s[i] of
          '>':
             begin
               if (_zz > '') or (_param <> '') then
                 D.isBad := true;
               if s > '' then
               begin
                 if not _start then
                 begin
                   if s > '' then
                     D.TagName := s;
                   D.RawTag := D.RawTag + s;
                 end else
                   D.isBad := true;
               end;
               D.RawTag := D.RawTag + _s[i];
             end;
          '"', '''':
             begin
               if _zz = '' then
               begin
                 D.isBad := true;
               end else
               begin
                 t := length(_zz);
                 if _zz[t] = '=' then
                 begin
                   _zz := _zz + _s[i];
                   if s > '' then
                     D.isBad := true;
                   D.RawTag := D.RawTag + _s[i];
                   s := '';
                 end else
                 begin
                   D.RawTag := D.RawTag + s + _s[i];
                   s := '';
                   _zz := '';
                   _param := '';
                 end;
               end;
             end;
          ' ', #13, #10, #9:
              begin
                if not _Start then
                begin
                  if s > '' then
                  begin
                    _Start := true;
                    D.RawTag := D.RawTag + s +_s[i];
                    D.TagName := s;
                    s := '';
                  end else
                    D.isBad := true;
                end else
                begin
                  if s <> '' then
                  begin
                    if _param = '' then
                    begin
                      _param := s;
                      D.RawTag := D.RawTag + s;
                      s := '';
                    end else
                    begin
                      if _zz = '' then
                        D.isBad := true;
                    end;
                  end;
                  D.RawTag := D.RawTag +_s[i];
                end;
              end;
          '<':
             begin
               if length(_zz) <> 2 then
                 D.isBad := true
               else
                 s := s + '&lt;';
             end;
          '&':
             begin
               if length(_zz) <> 2 then
                 D.isBad := true
               else
                 Correct_Entity(_S, i, s);
             end;
          '/':
             begin
               if not _Start then
               begin
                 D.RawTag := D.RawTag + '/';
                 if s = '' then
                 begin
                   D.isClose := true;
                 end else
                 begin
                   D.RawTag := D.RawTag + s;
                   D.isBad := true;
                   s := '';
                 end;
               end else
                 s := s + '/';
             end;
          '=':
             begin
               if _zz > '' then
                 D.isBad := true
               else
               begin
                  if (s > '') and (_param > '') then
                    D.isBad := true;
                 _zz := '=';
                 D.RawTag := D.RawTag + s + '=';
                 s := '';
               end;
             end;
          else
            s := s + _s[i];
        end;
        // если тэг запорчен
        if D.isBad then
        begin
          D.RawTag :=copy(_s, _num,  f - _num + 1);
          break;
        end;
      end;
      _num := f;
    end;
  end;

  procedure _BadTag(var s: string; var Stt: StackData);
  begin
    s := Stt.TextBeforeTag + CheckStrEntity(Stt.RawTag);
    SetLength(Stack, kol);
  end;

begin
  s := '';
  kol := 0;
  Corrected := '';
  UseXMLNS := false;
  _length := length(St);
  t := 0;
  i := 0;
  while i < _length do
  begin
    inc(i);
    case St[i] of
      '<':
         begin
           SetLength(Stack, kol + 1);
           Stack[kol].isBad := false;
           Stack[kol].isClose := false;
           Stack[kol].TagName := '';
           Stack[kol].RawTag := '<';
           Stack[kol].TextBeforeTag := s;
           s := '';
           _CheckTag(i, St, Stack[kol]);
           if Stack[kol].isBad then
             _BadTag(s, Stack[kol])
           else
           begin
             inc(t);
             if Stack[kol].isClose then
             begin
               if kol > 0 then
               begin
                 if Stack[kol].TagName = Stack[kol-1].TagName then
                 begin
                   dec(t);
                   if t <> 1 then
                   begin
                     Stack[kol-2].TextBeforeTag := Corrected + Stack[kol-2].TextBeforeTag;
                     Stack[kol-2].RawTag := Stack[kol-2].RawTag + Stack[kol-1].TextBeforeTag + Stack[kol-1].RawTag +
                                            Stack[kol].TextBeforeTag + Stack[kol].RawTag ;
                     Corrected := '';
                   end
                   else
                     Corrected := Corrected + Stack[kol-1].TextBeforeTag + Stack[kol-1].RawTag +
                                  Stack[kol].TextBeforeTag + Stack[kol].RawTag;
                   dec(t);
                   UseXMLNS := true;
                   dec(kol);
                   SetLength(Stack, kol);
                 end else
                  _BadTag(s, Stack[kol]);
               end else
                 _BadTag(s, Stack[kol]);
             end else
               inc(kol);
           end;
         end;
      '>': s := s + '&gt;';
      #13, #10:
         begin
           if St[i] = #13 then
           begin
             if i < _length then
             begin
               if St[i+1] = #10 then
                 s := s + '&#10;'
               else
                 s := s + St[i];
             end else
               s := s + St[i];
           end else
           begin
             if i > 1 then
             begin
               if St[i-1] <> #13 then
                 s := s + '&#10;'
               {else
                 s := s + St[i];}
             end else
               s := s + '&#10;';
           end;
         end;
      '&': Correct_Entity(St, i, s);
      '"': s := s + '&quot;';
      '''': s := s + '&apos;';
      else
        s := s + St[i];
    end;
  end;

  for i := 0 to kol - 1 do
    Corrected := Corrected + Stack[i].TextBeforeTag + CheckStrEntity(Stack[i].RawTag);

  Corrected := Corrected + s;
  SetLength(Stack, 0);
  Stack := nil;
end;

//переводит TColor в Hex RGB
function ColorToHTMLHex(Color: TColor): string;
var
  _RGB: integer;

begin
  _RGB := ColorToRGB(Color);
  //result := IntToHex(GetRValue(_RGB), 2) + IntToHex(GetGValue(_RGB), 2) + IntToHex(GetBValue(_RGB), 2);
  result := IntToHex(byte(_RGB), 2) + IntToHex(byte(_RGB shr 8), 2) + IntToHex(byte(_RGB shr 16), 2);
end;

//переводит Hex RGB (string) в TColor
function HTMLHexToColor(value: string): TColor;
var
  a: array [0..2] of integer;
  i, n, t: integer;

begin
  result := 0;
  if (value > '') then
  begin
    value := UpperCase(value);
    {$HINTS OFF}
    FillChar(a, sizeof(a), 0);
    {$HINTS ON}
    n := 0;
    if value[1] = '#' then delete(value, 1, 1);
    //А что, если будут цвета типа "black"?  {tut}
    for i := 1 to length(value) do
    begin
      if n > 2 then break;
      case value[i] of
        '0'..'9': t := ord(value[i]) - 48;
        'A'..'F': t := 10 + ord(value[i]) - 65;
        else
          t := 0;
      end;
      a[n] := a[n] * 16 + t;
      if i mod 2 = 0 then inc(n);
    end;
    result := a[2] shl 16 or a[1] shl 8 or a[0]//RGB(a[0], a[1], a[2]);
    //a[2] shl 16 or a[1] shl 8 or a[0]; = RGB
  end;
end;

//Перевести пиксели в типографский пункт (point)
//Input
//      inPixel: integer          - размер в пикселях
//      PixelSizeMM: real = 0.265 - размер пикселя
function PixelToPoint(inPixel: integer; PixelSizeMM: real = 0.265): real;
begin
  result := inPixel * PixelSizeMM / _PointToMM;
  //и оставим 2 знака после запятой ^_^
  result := round(result * 100) / 100;
end;

//Перевести типографский пункт (point) в пиксели
//Input
//      inPoint: integer          - размер в пикселях
//      PixelSizeMM: real = 0.265 - размер пикселя
function PointToPixel(inPoint: real; PixelSizeMM: real = 0.265): integer;
begin
  result := round(inPoint * _PointToMM / PixelSizeMM);
end;

//Перевести типографский пункт (point) в мм
//Input
//      inPoint: integer - размер в пунктах
function PointToMM(inPoint: real): real;
begin
  result := round(inPoint * _PointToMM * 100) / 100;
end;

//Перевести мм в типографский пункт (point)
//Input
//      inMM: integer - размер в пунктах
function MMToPoint(inMM: real): real;
begin
  result := round(inMM / _PointToMM * 100) / 100;
end;

//HorizontalAligment to Str
function HAlToStr(HA: TZHorizontalAlignment): string;
begin
  case HA of
    ZHAutomatic:              result := 'Automatic';
    ZHLeft:                   result := 'Left';
    ZHCenter:                 result := 'Center';
    ZHRight:                  result := 'Right';
    ZHFill:                   result := 'Fill';
    ZHJustify:                result := 'Justify';
    ZHCenterAcrossSelection:  result := 'CenterAcrossSelection';
    ZHDistributed:            result := 'Distributed';
    ZHJustifyDistributed:     result := 'JustifyDistributed';
  end;
end;

//Vertical Aligment to Str
function VAlToStr(VA: TZVerticalAlignment): string;
begin
  case VA of
    ZVAutomatic:          result := 'Automatic';
    ZVTop:                result := 'Top';
    ZVBottom:             result := 'Bottom';
    ZVCenter:             result := 'Center';
    ZVJustify:            result := 'Justify';
    ZVDistributed:        result := 'Distributed';
    ZVJustifyDistributed: result := 'JustifyDistributed';
  end;
end;

function ZBorderTypeToStr(ZB: TZBorderType): string;
begin
  case ZB of
    ZENone:         result := 'None';
    ZEContinuous:   result := 'Continuous';
    ZEDot:          result := 'Dot';
    ZEDash:         result := 'Dash';
    ZEDashDot:      result := 'DashDot';
    ZEDashDotDot:   result := 'DashDotDot';
    ZESlantDashDot: result := 'SlantDashDot';
    ZEDouble:       result := 'Double';
  end;
end;

function ZCellPatternToStr(pp: TZCellPattern): string;
begin
  case pp of
    ZPNone:           result := 'None';
    ZPSolid:          result := 'Solid';
    ZPGray75:         result := 'Gray75';
    ZPGray50:         result := 'Gray50';
    ZPGray25:         result := 'Gray25';
    ZPGray125:        result := 'Gray125';
    ZPGray0625:       result := 'Gray0625';
    ZPHorzStripe:     result := 'HorzStripe';
    ZPVertStripe:     result := 'VertStripe';
    ZPReverseDiagStripe: result := 'ReverseDiagStripe';
    ZPDiagStripe:     result := 'DiagStripe';
    ZPDiagCross:      result := 'DiagCross';
    ZPThickDiagCross: result := 'ThickDiagCross';
    ZPThinHorzStripe: result := 'ThinHorzStripe';
    ZPThinVertStripe: result := 'ThinVertStripe';
    ZPThinReverseDiagStripe: result := 'ThinReverseDiagStripe';
    ZPThinDiagStripe: result := 'ThinDiagStripe';
    ZPThinHorzCross:  result := 'ThinHorzCross';
    ZPThinDiagCross:  result := 'ThinDiagCross';
  end;
end;

function ZCellTypeToStr(pp: TZCellType): string;
begin
  case pp of
    ZENumber:         result := 'Number';
    ZEDateTime:       result := 'DateTime';
    ZEBoolean:        result := 'Boolean';
    ZEString:         result := 'String';
    ZEError:          result := 'Error';
  end;
end;

//Str to HorizontalAligment
function StrToHal(Value: string): TZHorizontalAlignment;
begin
  Value := UpperCase(Value);
  if Value = 'LEFT' then
    result := ZHLeft
  else if Value = 'CENTER' then
    result := ZHCenter
  else if Value = 'RIGHT' then
    result := ZHRight
  else if Value = 'FILL' then
    result := ZHFill
  else if Value = 'JUSTIFY' then
    result := ZHJustify
  else if Value = 'CENTERACROSSSELECTION' then
    result := ZHCenterAcrossSelection
  else if Value = 'DISTRIBUTED' then
    result := ZHDistributed
  else if Value = 'JUSTIFYDISTRIBUTED' then
    result := ZHJustifyDistributed
  else
    result := ZHAutomatic;
end;

//Str to Vertical Aligment
function StrToVAl(Value: string): TZVerticalAlignment;
begin
  Value := UpperCase(Value);
  if Value = 'CENTER' then
    result := ZVCenter
  else if Value = 'TOP' then
    result := ZVTop
  else if Value = 'BOTTOM' then
    result := ZVBottom
  else if Value = 'JUSTIFY' then
    result := ZVJustify
  else if Value = 'DISTRIBUTED' then
    result := ZVDistributed
  else if Value = 'JUSTIFYDISTRIBUTED' then
    result := ZVJustifyDistributed
  else
    result := ZVAutomatic;
end;

function StrToZBorderType(Value: string): TZBorderType;
begin
  Value := UpperCase(Value);
  if Value = 'CONTINUOUS' then
    result := ZEContinuous
  else if Value = 'DOT' then
    result := ZEDot
  else if Value = 'DASH' then
    result := ZEDash
  else if Value = 'DASHDOT' then
    result := ZEDashDot
  else if Value = 'DASHDOTDOT' then
    result := ZEDashDotDot
  else if Value = 'SLANTDASHDOT' then
    result := ZESlantDashDot
  else if Value = 'DOUBLE' then
    result := ZEDouble
  else
    result := ZENone;
end;

function StrToZCellPattern(Value: string): TZCellPattern;
begin
  Value := UpperCase(Value);
  if Value = 'SOLID' then
    result := ZPSolid
  else if Value = 'GRAY75' then
    result := ZPGray75
  else if Value = 'GRAY50' then
    result := ZPGray50
  else if Value = 'GRAY25' then
    result := ZPGray25
  else if Value = 'GRAY125' then
    result := ZPGray125
  else if Value = 'GRAY0625' then
    result := ZPGray0625
  else if Value = 'HORZSTRIPE' then
    result := ZPHorzStripe
  else if Value = 'VERTSTRIPE' then
    result := ZPVertStripe
  else if Value = 'REVERSEDIAGSTRIPE' then
    result := ZPReverseDiagStripe
  else if Value = 'DIAGSTRIPE' then
    result := ZPDiagStripe
  else if Value = 'DIAGCROSS' then
    result := ZPDiagCross
  else if Value = 'THICKDIAGCROSS' then
    result := ZPThickDiagCross
  else if Value = 'THINHORZSTRIPE' then
    result := ZPThinHorzStripe
  else if Value = 'THINVERTSTRIPE' then
    result := ZPThinVertStripe
  else if Value = 'THINREVERSEDIAGSTRIPE' then
    result := ZPThinReverseDiagStripe
  else if Value = 'THINDIAGSTRIPE' then
    result := ZPThinDiagStripe
  else if Value = 'THINHORZCROSS' then
    result := ZPThinHorzCross
  else if Value = 'THINDIAGCROSS' then
    result := ZPThinDiagCross
  else
    result := ZPNone;
end;

function StrToZCellType(Value: string): TZCellType;
begin
  Value := UpperCase(Value);
  if Value = 'NUMBER' then
    result := ZENumber
  else if Value = 'DATETIME' then
    result := ZEDateTime
  else if Value = 'BOOLEAN' then
    result := ZEBoolean
  else if Value = 'ERROR' then
    result := ZEError
  else
    result := ZEString;
end;

////::::::::::::: TZBorderStyle :::::::::::::::::////

constructor TZBorderStyle.Create();
begin
  FWeight := 0;
  FColor := clBlack;
  FLineStyle := ZENone;
end;

procedure TZBorderStyle.Assign(Source: TPersistent);
begin
  if Source is TZBorderStyle then
  begin
    Weight := (Source as TZBorderStyle).Weight;
    Color := (Source as TZBorderStyle).Color;
    LineStyle := (Source as TZBorderStyle).LineStyle;
  end else
    inherited Assign(Source);
end;

function TZBorderStyle.IsEqual(Source: TPersistent): boolean;
var
  zSource: TZBorderStyle;
begin
  Result := false;
  if not (Source is TZBorderStyle) then exit;
  zSource := Source as TZBorderStyle;

  if Self.LineStyle <> zSource.LineStyle then
    exit;

  if Self.Color <> zSource.Color then
    exit;

  if Self.Weight <> zSource.Weight then
    exit;

  Result := true;
end;

procedure TZBorderStyle.SetLineStyle(const Value: TZBorderType);
begin
  FLineStyle := Value;
end;

procedure TZBorderStyle.SetWeight(const Value: byte);
begin
  FWeight := Value;
  if FWeight > 3 then
    FWeight := 3;
end;

procedure TZBorderStyle.SetColor(const Value: TColor);
begin
  FColor := Value;
end;

////::::::::::::: TZBorder :::::::::::::::::////

constructor TZBorder.Create();
var
  i:integer;

begin
  for i := 0 to 5 do
    FBorder[i]:= TZBorderStyle.Create();
end;

destructor TZBorder.Destroy();
var
  i: integer;

begin
  for i := 0 to 5 do
    FreeAndNil(FBorder[i]);
  inherited Destroy;
end;

procedure TZBorder.Assign(Source: TPersistent);
var
  zSource: TZBorder;
  i: integer;

begin
  if (Source is TZBorder) then
  begin
    zSource := Source as TZBorder;
    for i := 0 to 5 do
      Border[i].Assign(zSource.Border[i]);
  end else
    inherited Assign(Source);
end;

function TZBorder.IsEqual(Source: TPersistent): boolean;
var
  zSource: TZBorder;
  i: integer;
begin
  Result := false;
  if not (Source is TZBorder) then exit;
  zSource := Source as TZBorder;

  for i := 0 to 5 do
  if not FBorder[i].IsEqual(zSource.Border[i]) then
    exit;

  Result := True;
end;

//установить стиль границы
procedure TZBorder.SetBorder(Num: integer; Const Value: TZBorderStyle);
begin
  if (Num>=0) and (Num <=3) then
    Border[num].Assign(Value);
end;

//прочитать стиль границы
function TZBorder.GetBorder(Num: integer): TZBorderStyle;
begin
  if (Num>=0) and (Num <=5) then
    result := FBorder[Num]
  else
    result := nil;
end;

////::::::::::::: TZAlignment :::::::::::::::::////

constructor TZAlignment.Create();
begin
  FHorizontal := ZHAutomatic;
  FIndent := 0;
  FRotate := 0;
  FShrinkToFit := false;
  FVertical := ZVAutomatic;
  FVerticalText := false;
  FWrapText := false;
end;

procedure TZAlignment.Assign(Source: TPersistent);
var
  zSource: TZAlignment;

begin
  if (Source is TZAlignment) then
  begin
    zSource := Source as TZAlignment;
    FHorizontal := zSource.Horizontal;
    FIndent := zSource.Indent;
    FRotate := zSource.Rotate;
    FShrinkToFit := zSource.ShrinkToFit;
    FVertical := zSource.Vertical;
    FVerticalText := zSource.VerticalText;
    FWrapText := zSource.WrapText;
  end else
    inherited Assign(Source);
end;

function TZAlignment.IsEqual(Source: TPersistent): boolean;
var
  zSource: TZAlignment;

begin
  result := true;
  if (Source is TZAlignment) then
  begin
    zSource := Source as TZAlignment;
    if Horizontal <> zSource.Horizontal then
    begin
      result := false;
      exit;
    end;
    if Indent <> zSource.Indent then
    begin
      result := false;
      exit;
    end;
    if Rotate <> zSource.Rotate then
    begin
      result := false;
      exit;
    end;
    if ShrinkToFit <> zSource.ShrinkToFit then
    begin
      result := false;
      exit;
    end;
    if Vertical <> zSource.Vertical then
    begin
      result := false;
      exit;
    end;
    if VerticalText <> zSource.VerticalText then
    begin
      result := false;
      exit;
    end;
    if WrapText <> zSource.WrapText then
    begin
      result := false;
      exit;
    end;
  end else
    result := false;
end;

procedure TZAlignment.SetHorizontal(const Value: TZHorizontalAlignment);
begin
  FHorizontal := Value;
end;

procedure TZAlignment.SetIndent(const Value: integer);
begin
  FIndent := Value;
end;

procedure TZAlignment.SetRotate(const Value: TZCellTextRotate);
begin
  FRotate := Value;
end;

procedure TZAlignment.SetShrinkToFit(const Value: boolean);
begin
  FShrinkToFit := Value;
end;

procedure TZAlignment.SetVertical(const Value: TZVerticalAlignment);
begin
  FVertical := Value;
end;

procedure TZAlignment.SetVerticalText(const Value:boolean);
begin
  FVerticalText := Value;
end;

procedure TZAlignment.SetWrapText(const Value: boolean);
begin
  FWrapText := Value;
end;

////::::::::::::: TZStyle :::::::::::::::::////
// about default font in Excel - http://support.microsoft.com/kb/214123
constructor TZStyle.Create();
begin
  FFont := TFont.Create();
  FFont.Size := 10;
  FFont.Name := 'Arial';
  FFont.Color := ClBlack;
  FBorder := TZBorder.Create();
  FAlignment := TZAlignment.Create();
  FBGColor := clWindow;
  FPatternColor := clWindow;
  FCellPattern := ZPNone;
  FNumberFormat := 'General';
  FProtect := true;
  FHideFormula := false;
end;

destructor TZStyle.Destroy();
begin
  FreeAndNil(FFont);
  FreeAndNil(FBorder);
  FreeAndNil(FAlignment);
  inherited Destroy();
end;

procedure TZStyle.Assign(Source: TPersistent);
var
  zSource: TZStyle;

begin
  if (Source is TZStyle) then
  begin
    zSource := Source as TZStyle;
    FFont.Assign(zSource.Font);
    FBorder.Assign(zSource.Border);
    FAlignment.Assign(zSource.Alignment);
    FBGColor := zSource.BGColor;
    FPatternColor := zSource.PatternColor;
    FCellPattern := zSource.CellPattern;
    FNumberFormat := zSource.NumberFormat;
    FProtect := zSource.Protect;
    FHideFormula := zSource.HideFormula;
  end else
    inherited Assign(Source);
end;

function TZStyle.IsEqual(Source: TPersistent): boolean;
var
  zSource: TZStyle;
begin
  Result := False;
  if not (Source is TZStyle) then exit;
  zSource := Source as TZStyle;

  if not Border.IsEqual(zSource.Border) then
    exit;

  if not self.Alignment.IsEqual(zSource.Alignment) then
    exit;

  if BGColor <> zSource.BGColor then
    exit;

  if PatternColor <> zSource.PatternColor then
    exit;

  if CellPattern <> zSource.CellPattern then
    exit;

  if NumberFormat <> zSource.NumberFormat then
    exit;

  if Protect <> zSource.Protect then
    exit;

  if HideFormula <> zSource.HideFormula then
    exit;

  //font
  if Font.Color <> zSource.Font.Color then
    exit;

  if Font.Height <> zSource.Font.Height then
    exit;

  if Font.Name <> zSource.Font.Name then
    exit;

  if Font.Pitch <> zSource.Font.Pitch then
    exit;

  if Font.Size <> zSource.Font.Size then
    exit;

  if Font.Style <> zSource.Font.Style then
    exit;
    {tut} //Ничего не забыл?

  Result := true;
end;

procedure TZStyle.SetFont(const Value: TFont);
begin
  FFont.Assign(Value);
end;

procedure TZStyle.SetBorder(const Value: TZBorder);
begin
  FBorder.Assign(Value);
end;

procedure TZStyle.SetAlignment(const Value: TZAlignment);
begin
  FAlignment.Assign(Value);
end;

procedure TZStyle.SetBGColor(const Value: TColor);
begin
  FBGColor := Value;
end;

procedure TZStyle.SetPatternColor(const Value: TColor);
begin
  FPatternColor := Value;
end;

procedure TZStyle.SetCellPattern(const Value: TZCellPattern);
begin
  FCellPattern := Value;
end;

procedure TZStyle.SetNumberFormat(const Value: string);
begin
  FNumberFormat := Value;
end;

////::::::::::::: TZStyles :::::::::::::::::////

constructor TZStyles.Create();
begin
  FDefaultStyle := TZStyle.Create();
  FCount := 0;
end;

destructor TZStyles.Destroy();
var
  i: integer;

begin
  FreeAndNil(FDefaultStyle);
  for i:= 0 to FCount - 1  do
    FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FStyles := nil;
  inherited Destroy();
end;

//Ищет стиль Style
//Возвращает:
//     -2          - стиль не найден
//     -1          - совпадает со стилем по умолчанию
//      0..Count-1 - номер стиля
function TZStyles.Find(const Style: TZStyle): integer;
var
  i: integer;

begin
  result := -2;
  if DefaultStyle.IsEqual(Style) then
  begin
    result := -1;
    exit;
  end;
  for i := 0 to Count - 1 do
  if Items[i].IsEqual(Style) then
  begin
    result := i;
    break;
  end;
end;

//Добавляет стиль
//      Style      - стиль
//      CheckMatch - проверять ли вхождение данного стиля
//Возвращает
//      Номер добавленного (или введённого ранее) стиля
function TZStyles.Add(const Style: TZStyle; CheckMatch: boolean = false): integer;
begin
  result := -2;
  if CheckMatch then
     result := Find(Style);
  if result = -2 then
  begin
    Count := Count + 1;
    Items[Count - 1].Assign(Style);
    result := Count - 1;
  end;
end;

procedure TZStyles.Assign(Source: TPersistent);
var
  zzz: TZStyles;
  i: integer;

begin
  if (Source is TZStyles) then
  begin
    zzz := Source as TZStyles;
    FDefaultStyle.Assign(zzz.DefaultStyle);
    Count := zzz.Count;
    for i := 0 to Count - 1 do
      FStyles[i].Assign(zzz[i]);
  end else
    inherited Assign(Source);
end;

procedure TZStyles.SetDefaultStyle(const Value: TZStyle);
begin
  if not(Value <> nil) then
    FDefaultStyle.Assign(Value);
end;

procedure TZStyles.SetCount(const Value: integer);
var
  i: integer;

begin
  if FCount < Value then
  begin
    SetLength(FStyles, Value);
    for i := FCount to Value - 1 do
    begin
      FStyles[i] := TZStyle.Create;
      FStyles[i].Assign(FDefaultStyle);
    end;
  end else
  begin
    for i:= Value to FCount - 1  do
      FreeAndNil(FStyles[i]);
    SetLength(FStyles, Value);
  end;
  FCount := Value;
end;

function TZStyles.GetStyle(num: integer): TZStyle;
begin
  if (num >= 0) and (num < Count) then
    result := FStyles[num]
  else
    result := DefaultStyle;
end;

procedure TZStyles.SetStyle(num: integer; const Value: TZStyle);
begin
  if (num >= 0) and (num < Count) then
    FStyles[num].Assign(Value)
  else
  if num = -1 then
    DefaultStyle.Assign(Value);
end;

//Удаляет стиль num, стили с большим номером сдвигаются
// num - номер стиля
//Возвращает:
//       0  - удалился успешно
//      -1  - стиль не удалился
function TZStyles.DeleteStyle(num: integer):integer;
begin
  if (num >= 0) and (num < Count) then
  begin
    FreeAndNil(FStyles[num]);
    System.Move(FStyles[num+1], FStyles[num], (Count - num - 1) * SizeOf(FStyles[num]));
    FStyles[Count - 1] := nil;
    dec(FCount);
    setlength(FStyles, FCount);
    // при удалении глянуть на ячейки - изменить num на 0, остальные сдвинуть.
    result := 0;
  end else
    result := -1;
end;

procedure TZStyles.Clear();
var
  i: integer;

begin
  //очистка стилей
  for i:= 0 to FCount - 1 do
    FreeAndNil(FStyles[i]);
  SetLength(FStyles,0);
  FCount := 0;
end;

////::::::::::::: TZCell :::::::::::::::::////

constructor TZCell.Create();
begin
  FFormula := '';
  FData := '';
  FHref := '';
  FComment := '';
  FCommentAuthor := '';
  FHRefScreenTip := '';
  FCellType := ZEString;
  FCellStyle := -1; //по дефолту
  FAlwaysShowComment := false;
  FShowComment := false;
end;

function TZCell.GetDataAsDouble: double;
Var err: integer;
begin
  Val(FData, Result, err); // need old-school to ignore regional settings
  if err>0 then Raise EConvertError.Create('ZxCell: Cannot cast data to number');
end;

function TZCell.GetDataAsInteger: integer;
begin
  Result := StrToInt(Data);
end;

procedure TZCell.SetDataAsInteger(const Value: integer);
begin
  Data := Trim(IntToStr(Value));
  CellType := ZENumber;
// Val adds the prepending space, maybe some I2S implementation would adds too
// and Excel dislikes it. Better safe than sorry.
end;

procedure TZCell.SetDataAsDouble(const Value: double);
var s: String; ss: ShortString;
begin
 if Value = 0
    then SetDataAsInteger(0) // work around Excel 2010 weird XLSX bug
    else begin
       Str(Value, ss); // need old-school to ignore regional settings
       s := string(ss); // make XE2 happy

       s := UpperCase(Trim(s));
      // UpperCase for exponent w.r.t OpenXML format
      // Trim for leading space w.r.t XML SS format

       FData := s;

       CellType := ZENumber;
      // Seem natural and logical thing to do w.r.t further export...
      // Seem out of "brain-dead no-automation overall aproach of a component...
      // Correct choice? dunno. I prefer making export better
    end
end;

procedure TZCell.Assign(Source: TPersistent);
var
  zSource: TZCell;

begin
  if (Source is TZCell) then
  begin
    zSource := Source as TZCell;
    FFormula := zSource.Formula;
    FData := zSource.Data;
    FHref := zSource.Href;
    FComment := zSource.Comment;
    FCommentAuthor := zSource.CommentAuthor;
    FCellStyle := zSource.CellStyle;
    FCellType := zSource.CellType;
    FAlwaysShowComment := zSource.AlwaysShowComment;
    FShowComment := zSource.ShowComment;
  end else
    inherited Assign(Source);
end;

//Очистка ячейки
procedure TZCell.Clear();
begin
  FFormula := '';
  FData := '';
  FHref := '';
  FComment := '';
  FCommentAuthor := '';
  FHRefScreenTip := '';
  FCellType := ZEString;
  FCellStyle := -1;
  FAlwaysShowComment := false;
  FShowComment := false;
end;

////::::::::::::: TZMergeCells :::::::::::::::::////

constructor TZMergeCells.Create(ASheet: TZSheet);
begin
  FSheet := ASheet;
  FCount := 0;
end;

destructor TZMergeCells.Destroy();
begin
  Clear();
  inherited Destroy();
end;

procedure TZMergeCells.Clear();
begin
  SetLength(FMergeArea, 0);
  FCount := 0;
end;

function TZMergeCells.GetItem(num: integer):TRect;
begin
  if (num >= 0) and (num < Count) then
    Result := FMergeArea[num]
  else
  begin
    result.Left   := 0;
    result.Top    := 0;
    result.Right  := 0;
    result.Bottom := 0;
  end;
end;

//Объединить ячейки входящие в прямоуголтник
//принимает прямоугольник (Rct: TRect)
//возвращает:
//      0 - всё нормально, ячейки объединены
//      1 - указанный прямоугольник выходит за границы, ячейки не объединены
//      2 - указанный прямоуголник пересекается(входит) введённые ранее, ячейки не объединены
//      3 - прямоугольник из одной ячейки не добавляет
function TZMergeCells.AddRect(Rct:TRect): byte;
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
  if Rct.Left > Rct.Right then
  begin
    i := Rct.Left;
    Rct.Left := Rct.Right;
    Rct.Right := i;
  end;
  if Rct.Top > Rct.Bottom then
  begin
    i := Rct.Top;
    Rct.Top := Rct.Bottom;
    Rct.Bottom := i;
  end;
  if (Rct.Left < 0) or (Rct.Top < 0) then
  begin
    result := 1;
    exit;
  end;
  if (Rct.Right - Rct.Left = 0) and (Rct.Bottom - Rct.Top = 0) then
  begin
    result := 3;
    exit;
  end;
  for i := 0 to Count-1 do
  {if (((FMergeArea[i].Left >= Rct.Left) and (FMergeArea[i].Left <= Rct.Right)) or
      ((FMergeArea[i].right >= Rct.Left) and (FMergeArea[i].right <= Rct.Right))) and
     (((FMergeArea[i].Top >= Rct.Top) and (FMergeArea[i].Top <= Rct.Bottom)) or
      ((FMergeArea[i].Bottom >= Rct.Top) and (FMergeArea[i].Bottom <= Rct.Bottom))) then}
  if usl(FMergeArea[i], Rct) or usl(rct, FMergeArea[i]) then 
  begin
    result := 2;
    exit;
  end;
  //если надо, увеличиваем кол-во строк/столбцов в хранилище
  if (FSheet <> nil) then
  begin
    if Rct.Right > FSheet.ColCount - 1 then FSheet.ColCount := Rct.Right {+ 1};
    if Rct.Bottom > FSheet.RowCount - 1 then FSheet.RowCount := Rct.Bottom {+ 1};  
  end;
  inc(FCount);
  setlength(FMergeArea, FCount);
  FMergeArea[FCount - 1] := Rct;
  result := 0;
end;

//Проверка на вхождение ячейки в левый верхний угол области (
//принимает (ACol, ARow: integer) -  координаты ячейки
//возвращает
//      -1              ячейка не является левым верхним углом области
//      int num >=0     номер области, в которой ячейка есть левый верхний угол
function TZMergeCells.InLeftTopCorner(ACol, ARow: integer): integer;
var
  i: integer;

begin
  result := -1;
  for i := 0 to FCount - 1 do
  if (ACol = FMergeArea[i].Left) and (ARow = FMergeArea[i].top) then
  begin
    result := i;
    break;
  end;
end;

//Проверка на вхождение ячейки в область
//принимает (ACol, ARow: integer) -  координаты ячейки
//возвращает
//      -1              ячейка не входит в область
//      int num >=0     номер области, в которой содержится ячейка
function TZMergeCells.InMergeRange(ACol, ARow: integer): integer;
var
  i: integer;

begin
  result := -1;
  for i := 0 to FCount - 1 do
  if (ACol >= FMergeArea[i].Left) and (ACol <= FMergeArea[i].Right) and
     (ARow >= FMergeArea[i].Top) and (ARow <= FMergeArea[i].Bottom) then
  begin
    result := i;
    break;
  end;
end;

//удаляет облать num
//возвращает
//      - TRUE - область удалена
//      - FALSE - num>Count-1 или num<0
function TZMergeCells.DeleteItem(num: integer): boolean;
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

//Объединить ячейки входящие в прямоуголтник
//принимает прямоугольник (x1, y1, x2, y2: integer)
//возвращает:
//      0 - всё нормально, ячейки объединены
//      1 - указанный прямоугольник выходит за границы, ячейки не объединены
//      2 - указанный прямоуголник пересекается(входит) введённые ранее, ячейки не объединены
//      3 - прямоугольник из одной ячейки не добавляет
function TZMergeCells.AddRectXY(x1,y1,x2,y2: integer): byte;
var
  Rct: TRect;
begin
  Rct.left := x1;
  Rct.Top := y1;
  Rct.Right := x2;
  Rct.Bottom := y2;
  result := AddRect(Rct);
end;

////::::::::::::: TZRowColOptions :::::::::::::::::////

constructor TZRowColOptions.Create(ASheet: TZSheet);
begin
  inherited Create();
  FSheet  := ASheet;
  FHidden := false;
  FAuto  := true;
  FStyleID := -1;
  FBreaked := false;
end;

procedure TZRowColOptions.Assign(Source: TPersistent);
begin
  if Source is TZRowColOptions then
  begin
    Hidden := (Source as TZRowColOptions).Hidden;
    StyleID := (Source as TZRowColOptions).StyleID;
    FSize := (Source as TZRowColOptions).FSize;
    FAuto := (Source as TZRowColOptions).FAuto;
    FBreaked := (Source as TZRowColOptions).Breaked;
  end else
    inherited Assign(Source);
end;

function  TZRowColOptions.GetAuto(): boolean;
begin
  result := FAuto;
end;

procedure TZRowColOptions.SetAuto(Value: boolean);
begin
  FAuto := Value;
end;

function  TZRowColOptions.GetSizePoint(): real;
begin
  result := FSize;
end;

procedure TZRowColOptions.SetSizePoint(Value: real);
begin
  if Value < 0 then  exit;
  FSize := Value;
end;

function  TZRowColOptions.GetSizeMM(): real;
begin
  result := PointToMM(FSize);
end;

procedure TZRowColOptions.SetSizeMM(Value: real);
begin
  FSize := MMToPoint(Value);
end;

////::::::::::::: TZRowOptions :::::::::::::::::////

constructor TZRowOptions.Create(ASheet: TZSheet);
begin
  inherited Create(ASheet);
  FSize := 48;
end;

function TZRowOptions.GetSizePix(): integer;
var
  t :real;

begin
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.HorPixelSize;
  result := PointToPixel(FSize, t);
end;

procedure TZRowOptions.SetSizePix(Value: integer);
var
  t :real;

begin
  if Value < 0 then  exit;
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.HorPixelSize;
  FSize := PixelToPoint(Value, t);
end;

////::::::::::::: TZColOptions :::::::::::::::::////

constructor TZColOptions.Create(ASheet: TZSheet);
begin
  inherited Create(ASheet);
  FSize := 12.75;
end;

function TZColOptions.GetSizePix(): integer;
var
  t :real;

begin
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.VertPixelSize;
  result := PointToPixel(FSize, t);
end;

procedure TZColOptions.SetSizePix(Value: integer);
var
  t :real;

begin
  if Value < 0 then  exit;
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.VertPixelSize;
  FSize := PixelToPoint(Value, t);
end;

////::::::::::::: TZSheetFooterHeader :::::::::::::::::////

constructor TZSheetFooterHeader.Create();
begin
  inherited;
  FDataLeft := '';
  FData := '';
  FDataRight := '';
end;

procedure TZSheetFooterHeader.Assign(Source: TPersistent);
var
  _t: TZSheetFooterHeader;

begin
  if (Source is TZSheetFooterHeader) then
  begin
    _t := Source as TZSheetFooterHeader;
    FDataLeft := _t.DataLeft;
    FData := _t.Data;
    FDataRight := _t.DataRight;
  end
  else
    inherited;
end;

function TZSheetFooterHeader.IsEqual(Source: TPersistent): boolean;
var
  _t: TZSheetFooterHeader;

begin
  result := false;
  if (Source is TZSheetFooterHeader) then
  begin
    _t := Source as TZSheetFooterHeader;
    result := (_t.DataLeft = FDataLeft) and
              (_t.Data = FData) and
              (_t.DataRight = FDataRight);
  end;
end;

////::::::::::::: TZSheetOptions :::::::::::::::::////

constructor TZSheetOptions.Create();
begin
  inherited;
  FActiveCol := 0;
  FActiveRow := 0;
  FMarginBottom := 25;
  FMarginLeft := 20;
  FMarginTop := 25;
  FMarginRight := 20;
  FPortraitOrientation := true;
  FCenterHorizontal := false;
  FCenterVertical := false;
  FStartPageNumber := 1 ;
  FHeaderMargin := 13;
  FFooterMargin := 13;
  FPaperSize := 9;

  FIsEvenFooterEqual := true;
  FIsEvenHeaderEqual := true;

  FHeader := TZSheetFooterHeader.Create();
  FFooter := TZSheetFooterHeader.Create();
  FEvenHeader := TZSheetFooterHeader.Create();
  FEvenFooter := TZSheetFooterHeader.Create();

  FHeaderBGColor := clWindow;
  FFooterBGColor := clWindow;

  FSplitVerticalMode := ZSplitNone;
  FSplitHorizontalMode := ZSplitNone;
  FSplitVerticalValue := 0;
  FSplitHorizontalValue := 0;
  FPaperWidth := 0;
  FPaperHeight := 0;
end;

destructor TZSheetOptions.Destroy();
begin
  FreeAndNil(FHeader);
  FreeAndNil(FFooter);
  FreeAndNil(FEvenHeader);
  FreeAndNil(FEvenFooter);
  inherited;
end;

function TZSheetOptions.GetHeaderData(): string;
begin
  result := FHeader.Data;
end;

procedure TZSheetOptions.SetHeaderData(Value: string);
begin
  FHeader.Data := Value;
end;

function TZSheetOptions.GetFooterData(): string;
begin
  result := FFooter.Data;
end;

procedure TZSheetOptions.SetfooterData(Value: string);
begin
  FFooter.Data := Value;
end;

procedure TZSheetOptions.Assign(Source: TPersistent);
var
  t: TZSheetOptions;

begin
  if Source is TZSheetOptions then
  begin
    t := Source as TZSheetOptions;
    ActiveCol := t.ActiveCol;
    ActiveRow := t.ActiveRow;
    MarginBottom := t.MarginBottom;
    MarginLeft := t.MarginLeft;
    MarginTop := t.MarginTop;
    MarginRight := t.MarginRight;
    PortraitOrientation := t.PortraitOrientation;
    CenterHorizontal := t.CenterHorizontal;
    CenterVertical := t.CenterVertical;
    StartPageNumber := t.StartPageNumber;
    HeaderMargin := t.HeaderMargin;
    FooterMargin := t.FooterMargin;
    HeaderData := t.HeaderData;
    FooterData := t.FooterData;
    PaperSize := t.PaperSize;
    SplitVerticalMode := t.SplitVerticalMode;
    SplitHorizontalMode := t.SplitHorizontalMode;
    SplitVerticalValue := t.SplitVerticalValue;
    SplitHorizontalValue := t.SplitHorizontalValue;
    Footer.Assign(t.Footer);
    Header.Assign(t.Header);
    EvenHeader.Assign(t.EvenHeader);
    EvenFooter.Assign(t.EvenFooter);
    HeaderBGColor := t.HeaderBGColor;
    FooterBGColor := t.FooterBGColor;
    IsEvenFooterEqual := t.IsEvenFooterEqual;
    IsEvenHeaderEqual := t.IsEvenHeaderEqual;
  end else
    inherited Assign(Source);
end;

////::::::::::::: TZSheet :::::::::::::::::////

constructor TZSheet.Create(AStore: TZEXMLSS);
var
 i, j: integer;

begin
  FStore := AStore;
  FRowCount := 0;
  FColCount := 0;
  FDefaultRowHeight := 12.75;//16;
  FDefaultColWidth := 48;//60;
  FMergeCells := TZMergeCells.Create(self);
  setlength(FCells, FColCount);
  FTabColor := ClWindow;
  FProtect := false;
  FRightToLeft := false;
  FSelected := false;
  SetLength(FRows, FRowCount);
  SetLength(FColumns, FColCount);
  for i := 0 to FColCount - 1 do
  begin
    setlength(FCells[i], FRowCount);
    for j := 0 to FRowCount - 1 do
    begin
      FCells[i,j] := TZCell.Create();
    end;
    FColumns[i] := TZColOptions.create(self);
    FColumns[i].Width := DefaultColWidth;
  end;
  for i := 0 to FRowCount - 1 do
  begin
    FRows[i] := TZRowOptions.create(self);
    FRows[i].Height := DefaultRowHeight;
  end;
  FSheetOptions := TZSheetOptions.Create();
  if FStore <> nil then
    if FStore.DefaultSheetOptions <> nil then
      FSheetOptions.Assign(FStore.DefaultSheetOptions);

  FPrintRows := TZSheetPrintTitles.Create(Self, false);
  FPrintCols := TZSheetPrintTitles.Create(Self, true);
  {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
  FConditionalFormatting := TZConditionalFormatting.Create();
  {$ENDIF}
end;

destructor TZSheet.Destroy();
begin
  try
    FreeAndNil(FMergeCells);
    FreeAndNil(FSheetOptions);
    FPrintRows.Free;
    FPrintCols.Free;
    Clear();
    FCells := nil;
    FRows := nil;
    FColumns := nil;
    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    FreeAndNil(FConditionalFormatting);
    {$ENDIF}
  finally
    inherited Destroy();
  end;
end;

procedure TZSheet.Assign(Source: TPersistent);
var
  zSource: TZSheet;
  i, j: integer;

begin
  if (Source is TZSheet) then
  begin
    ZSource     := Source as TZSheet;
    ////////////////////////////////
    RowCount    := zSource.RowCount;
    ColCount    := zSource.ColCount;
    TabColor    := zSource.TabColor;
    Title       := zSource.Title;
    Protect     := zSource.Protect;
    RightToLeft := zSource.RightToLeft;
    DefaultRowHeight := zSource.DefaultRowHeight;
    DefaultColWidth := zSource.DefaultColWidth;

    for i := 0 to RowCount - 1 do
      Rows[i] := ZSource.Rows[i];

    for i := 0 to ColCount - 1 do
    begin
      Columns[i] := ZSource.Columns[i];
      for j := 0 to RowCount - 1 do
        Cell[i, j] := zSource.Cell[i, j];
    end;

    FSelected := zSource.Selected;

    SheetOptions.Assign(zSource.SheetOptions);

    //Объединённые ячейки
    MergeCells.Clear();
    for i := 0 to ZSource.MergeCells.Count - 1 do
      MergeCells.AddRect(ZSource.MergeCells.GetItem(i));

    {дописать что ещё нужно копировать}
    {tut!}

    {$IFDEF ZUSE_CONDITIONAL_FORMATTING}
    ConditionalFormatting.Assign(zSource.ConditionalFormatting);
    {$ENDIF}

    RowsToRepeat.Assign(zSource.RowsToRepeat);
    ColsToRepeat.Assign(zSource.ColsToRepeat);
  end else
    inherited Assign(Source);
end;

{$IFDEF ZUSE_CONDITIONAL_FORMATTING}
procedure TZSheet.SetConditionalFormatting(Value: TZConditionalFormatting);
begin
  if (Assigned(Value)) then
    FConditionalFormatting.Assign(Value);
end; //SetConditionalFormatting
{$ENDIF}

function TZSheet.GetSheetOptions(): TZSheetOptions;
begin
  result := FSheetOptions;
end;

procedure TZSheet.SetSheetOptions(Value: TZSheetOptions);
begin
  if Value <> nil then
   FSheetOptions.Assign(Value);
end;

procedure TZSheet.SetColumn(num: integer; const Value:TZColOptions);
begin
  if (num >= 0) and (num < FColCount) then
    FColumns[num].Assign(Value);
end;

function  TZSheet.GetColumn(num: integer): TZColOptions;
begin
  if (num >= 0) and (num < FColCount) then
    result := FColumns[num]
  else
    result := nil;
end;

procedure TZSheet.SetRow(num: integer; const Value:TZRowOptions);
begin
  if (num >= 0) and (num < FRowCount) then
    FRows[num].Assign(Value);
end;

function  TZSheet.GetRow(num: integer): TZRowOptions;
begin
  if (num >= 0) and (num < FRowCount) then
    result := FRows[num]
  else
    result := nil;
end;

procedure TZSheet.SetColWidth(num: integer; const Value: real);
begin
  if (num < ColCount) and (num >= 0) and (Value >= 0) then
  if FColumns[num] <> nil then
    FColumns[num].Width := Value;
end;

function TZSheet.GetColWidth(num: integer): real;
begin
  result := 0;
  if (num < ColCount) and (num >= 0) then
    if FColumns[num] <> nil then
      result := FColumns[num].Width
end;

procedure TZSheet.SetRowHeight(num: integer; const Value: real);
begin
  if (num < RowCount) and (num >= 0) and (Value >= 0) then
    if FRows[num] <> nil then
      FRows[num].Height := Value;
end;

function TZSheet.GetRowHeight(num: integer): real;
begin
  result := 0;
  if (num < RowCount) and (num >= 0) then
    if FRows[num] <> nil then
      result := FRows[num].Height
end;

//установить ширину столбца по умолчанию
procedure TZSheet.SetDefaultColWidth(const Value: real);
begin
  if Value >= 0 then
    FDefaultColWidth := round(Value*100)/100;
end;

//установить высоту строки по умолчанию
procedure TZSheet.SetDefaultRowHeight(const Value: real);
begin
  if Value >= 0 then
    FDefaultRowHeight := round(Value*100)/100;
end;

procedure TZSheet.SetPrintCols(const Value: TZSheetPrintTitles);
begin
  FPrintCols.Assign( Value );
end;

procedure TZSheet.SetPrintRows(const Value: TZSheetPrintTitles);
begin
  FPrintRows.Assign( Value );
end;

procedure TZSheet.Clear();
var
  i, j: integer;

begin
  //очистка ячеек
  for i:= 0 to FColCount - 1 do
  begin
    for j:= 0 to FRowCount - 1 do
      FreeAndNil(FCells[i][j]);
    SetLength(FCells[i], 0);
    FColumns[i].Free;
    FCells[i] := nil;
  end;
  for i := 0 to FRowCount - 1 do
    FRows[i].Free;
  SetLength(FCells, 0);
  FRowCount := 0;
  FColCount := 0;
  SetLength(FRows, 0);
  SetLength(FColumns, 0);
end;

procedure TZSheet.SetCell(ACol, ARow: integer; const Value: TZCell);
begin
  //добавить проверку на корректность текста
  if (ACol >= 0) and (ACol < FColCount) and
     (ARow >= 0) and (ARow < FRowCount) then
  FCells[ACol, ARow].Assign(Value);
end;

function TZSheet.GetCell(ACol, ARow: integer):TZCell;
begin
  if (ACol >= 0) and (ACol < FColCount) and
     (ARow >= 0) and (ARow < FRowCount) then
  result := FCells[ACol, ARow]
    else
  result := nil;
end;

//установить кол-во столбцов
procedure TZSheet.SetColCount(const Value: integer);
var
  i, j: integer;

begin
  if Value < 0 then exit;
  if FColCount > Value then // todo Repeatable columns may be affected
  begin
    for i := Value to FColCount - 1 do
    begin
      for j := 0 to FRowCount - 1 do
        FreeAndNil(FCells[i][j]);
      setlength(FCells[i], 0);
      FreeAndNil(FColumns[i]);
    end;
    setlength(FCells, Value);
    SetLength(FColumns, Value);
  end else
  begin
    setlength(FCells, Value);
    SetLength(FColumns, Value);
    for i := FColCount to Value - 1 do
    begin
      setlength(FCells[i], FRowCount);
      FColumns[i] := TZColOptions.Create(self);
      FColumns[i].Width := DefaultColWidth;
      for j := 0 to FRowCount - 1 do
        FCells[i][j] := TZCell.Create();
    end;
  end;
  FColCount := Value;
end;

//получить кол-во столбцов
function TZSheet.GetColCount: integer;
begin
  result := FColCount;
end;

//установить кол-во строк
procedure TZSheet.SetRowCount(const Value: integer);
var
  i, j: integer;

begin
  if Value < 0 then exit;
  if FRowCount > Value then  // todo Repeatable rows may be affected
  begin
    for i := 0 to FColCount - 1 do
    begin
      for j := Value  to FRowCount - 1 do
        FreeAndNil(FCells[i][j]);
      setlength(FCells[i], Value);
    end;
    for i := Value to FRowCount - 1 do
      FreeAndNil(FRows[i]);
    SetLength(FRows, Value);
  end else
  begin
    for i := 0 to FColCount - 1 do
    begin
      setlength(FCells[i], Value);
      for j := FRowCount to Value - 1 do
        FCells[i][j] := TZCell.Create();
    end;
    SetLength(FRows, Value);
    for i := FRowCount to Value - 1 do
    begin
      FRows[i] := TZRowOptions.Create(self);
      FRows[i].Height := DefaultRowHeight;
    end;
  end;
  FRowCount := Value;
end;

//получить кол-во строк
function TZSheet.GetRowCount: integer;
begin
  result := FRowCount;
end;


////::::::::::::: TZSheets :::::::::::::::::////

constructor TZSheets.Create(AStore: TZEXMLSS);
begin
  FStore := AStore;
  FCount := 0;
  SetLength(FSheets, 0);
end;

destructor TZSheets.Destroy();
var
  i: integer;

begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FSheets[i]);
  SetLength(FSheets, 0);
  FSheets := nil;
  inherited Destroy();
end;

procedure TZSheets.Assign(Source: TPersistent);
var
  t: TZSheets;
  i: integer;

begin
  if (Source is TZSheets) then
  begin
    t := Source as TZSheets;
    Count := t.Count;
    for i := 0 to Count - 1 do
      Sheet[i].Assign(t.Sheet[i]);
  end else
    inherited Assign(Source);
end; //Assign

procedure TZSheets.SetSheetCount(const Value: integer);
var
  i: integer;

begin
  if Count < Value then
  begin
    SetLength(FSheets, Value);
    for i := Count to Value - 1 do
      FSheets[i] := TZSheet.Create(FStore);
  end else
  begin
    for i:= Value to Count - 1  do
      FreeAndNil(FSheets[i]);
    SetLength(FSheets, Value);
  end;
  FCount := Value;
end;

procedure TZSheets.SetSheet(num: integer; Const Value: TZSheet);
begin
  if (num < Count) and (num >= 0) then
    FSheets[num].Assign(Value);
end;

function TZSheets.GetSheet(num: Integer): TZSheet;
begin
  if (num < Count) and (num >= 0) then
    result := FSheets[num]
  else
    result := nil;
end;

////::::::::::::: TZEXMLDocumentProperties :::::::::::::::::////

constructor TZEXMLDocumentProperties.Create();
begin
  FAuthor := 'none';
  FLastAuthor := 'none';
  FCompany := 'none';
  FVersion := '11.9999';
  FCreated := date + time;
  FWindowHeight := 20000;
  FWindowWidth := 20000;
  FWindowTopX := 150;
  FWindowTopY := 150;
  FModeR1C1 := false;
end;

procedure TZEXMLDocumentProperties.SetAuthor(const Value: string);
begin
  //поставить проверку на валидность ввода
  FAuthor := Value;
end;

procedure TZEXMLDocumentProperties.SetLastAuthor(const Value: string);
begin
  //поставить проверку на валидность ввода
  FLastAuthor := Value;
end;

procedure TZEXMLDocumentProperties.SetCompany(const Value: string);
begin
  //поставить проверку на валидность ввода
  FCompany := Value;
end;

procedure TZEXMLDocumentProperties.SetVersion(const Value: string);
begin
  //поставить проверку на валидность ввода
  FVersion := Value;
end;

procedure TZEXMLDocumentProperties.Assign(Source: TPersistent);
var
  zzz: TZEXMLDocumentProperties;

begin
  if Source is TZEXMLDocumentProperties then
  begin
    zzz := Source as TZEXMLDocumentProperties;
    Author := zzz.Author;
    LastAuthor := zzz.LastAuthor;
    Created := zzz.Created;
    Company := zzz.Company;
    Version := zzz.Version;
    WindowHeight := zzz.WindowHeight;
    WindowWidth := zzz.WindowWidth;
    WindowTopX := zzz.WindowTopX;
    WindowTopY := zzz.WindowTopY;
    ModeR1C1 := zzz.ModeR1C1;
  end else
    inherited Assign(Source);
end;

////::::::::::::: TZEXMLSS :::::::::::::::::////

constructor TZEXMLSS.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDocumentProperties := TZEXMLDocumentProperties.Create;
  FStyles := TZStyles.Create();
  FSheets := TZSheets.Create(self);
  FHorPixelSize := 0.265;
  FVertPixelSize := 0.265;
  FDefaultSheetOptions := TZSheetOptions.Create();
end;

destructor TZEXMLSS.Destroy();
begin
  FreeAndNil(FDefaultSheetOptions);
  FreeAndNil(FDocumentProperties);
  FreeAndNil(FStyles);
  FreeAndNil(FSheets);
  inherited Destroy();
end;

procedure TZEXMLSS.Assign(Source: TPersistent);
var
  t: TZEXMLSS;

begin
  if (Source is TZEXMLSS) then
  begin
    t := Source as TZEXMLSS;
    Styles.Assign(t.Styles);
    Sheets.Assign(t.Sheets);
  end else
  if (Source is TZStyles) then
    Styles.Assign(Source as TZStyles)
  else
  if (Source is TZSheets) then
    Sheets.Assign(Source as TZSheets)
  else
    inherited Assign(Source);
end; //Assign

procedure TZEXMLSS.SetHorPixelSize(Value: real);
begin
  if Value > 0 then
    FHorPixelSize := Value;
end;

procedure TZEXMLSS.SetVertPixelSize(Value: real);
begin
  if Value > 0 then
    FVertPixelSize := Value;
end;

//Получить HorPixelSize и VertPixelSize устройства hdc
procedure TZEXMLSS.GetPixelSize(hdc: HWND);
begin
  {$IFDEF FPC}
  {$ELSE}
  HorPixelSize  := GetDeviceCaps(hdc, HORZSIZE) / GetDeviceCaps(hdc, HORZRES); //горизонтальный размер пикселя в миллиметрах
  VertPixelSize := GetDeviceCaps(hdc, VERTSIZE) / GetDeviceCaps(hdc, VERTRES); //вертикальный размер пикселя в миллиметрах
  {$ENDIF}
end;

function TZEXMLSS.GetDefaultSheetOptions(): TZSheetOptions;
begin
  result := FDefaultSheetOptions;
end;

procedure TZEXMLSS.SetDefaultSheetOptions(Value: TZSheetOptions);
begin
  if Value <> nil then
   FDefaultSheetOptions.Assign(Value);
end;


//Сохраняет в файл FileName
//      SheetsNumbers: array of integer  - массив номеров страниц в нужной последовательности
//      SheetsNames: array of ansistring     - массив названий страниц
//              количество элементов в двух массивах должны совпадать
//      CodePage: byte                   - кодировка:
//                                         0 - UTF-8
//                                         1 - windows-1251
//function TZEXMLSS.SaveToFile(const FileName: ansistring; const SheetsNumbers:array of integer;const SheetsNames: array of ansistring; CodePage: byte{$IFDEF VER130}{$ELSE} = 0{$ENDIF}): integer;
{var
  Stream: TStream;

begin
  result := 0;
  try
    try
      Stream := TFileStream.Create(FileName, fmCreate);
    except
      result := 1;
    end;
    if result = 0 then
      result := SaveToStream(Stream, SheetsNumbers, SheetsNames, CodePage);
  finally
    if Stream <> nil then
      Stream.Free;
  end;
end;
}

//Сохраняет в поток
//      SheetsNumbers: array of integer  - массив номеров страниц в нужной последовательности
//      SheetsNames: array of ansistring     - массив названий страниц
//              количество элементов в двух массивах должны совпадать
//      CodePage: byte                   - кодировка:
//                                         0 - UTF-8
//                                         1 - windows-1251
//function TZEXMLSS.SaveToStream(Stream: TStream; const SheetsNumbers:array of integer;const SheetsNames: array of ansistring; CodePage: byte {$IFDEF VER130}{$ELSE} = 0{$ENDIF}): integer;
{begin
end;}

//Сохраняет в файл FileName
//      SheetsNumbers: array of integer  - массив номеров страниц в нужной последовательности
//      CodePage: byte                   - кодировка:
//                                         0 - UTF-8
//                                         1 - windows-1251
{function TZEXMLSS.SaveToFile(const FileName: ansistring; const SheetsNumbers:array of integer; CodePage: byte = 0): integer;
begin
  result := SaveToFile(FileName, SheetsNumbers, [], CodePage);
end;}

//Сохраняет в поток
//      SheetsNumbers: array of integer  - массив номеров страниц в нужной последовательности
//      CodePage: byte                   - кодировка:
//                                         0 - UTF-8
//                                         1 - windows-1251
{function TZEXMLSS.SaveToStream(Stream: TStream; const SheetsNumbers:array of integer; CodePage: byte = 0): integer;
begin
  result := SaveToStream(Stream, SheetsNumbers, [], CodePage);
end;}

//Сохраняет в файл FileName
//      CodePage: byte                   - кодировка:
//                                         0 - UTF-8
//                                         1 - windows-1251
{function TZEXMLSS.SaveToFile(const FileName: ansistring; CodePage: byte = 0): integer;
begin
  result := SaveToFile(FileName, [], [], CodePage);
end;}

//Сохраняет в поток
//      CodePage: byte                   - кодировка:
//                                         0 - UTF-8
//                                         1 - windows-1251
{function TZEXMLSS.SaveToStream(Stream: TStream; CodePage: byte = 0): integer;
begin
  result := SaveToStream(Stream, [], [], CodePage);
end;}

//регистрируем компонент
procedure Register();
begin
  RegisterComponents('ZColor', [TZEXMLSS]);
end;

{ TZSheetPrintTitles }

procedure TZSheetPrintTitles.Assign(Source: TPersistent);
var f, t: word;  a: boolean;
begin
  if Source is TZSheetPrintTitles then
    begin
      F   := TZSheetPrintTitles(Source).From;
      T   := TZSheetPrintTitles(Source).Till;
      A   := TZSheetPrintTitles(Source).Active;

      if A then RequireValid(F, T);

      FFrom   := F;
      FTill   := T;
      FActive := A;
    end
  else inherited;
end;

constructor TZSheetPrintTitles.Create(const owner: TZSheet; const ForColumns: boolean);
begin
  if nil = owner then raise Exception.Create(Self.ClassName+' requires an existing worksheet for owners.');

  Self.FOwner := owner;
  Self.FColumns := ForColumns;
end;

procedure TZSheetPrintTitles.SetActive(const Value: boolean);
begin
  if Value then RequireValid(FFrom, FTill);

  FActive := Value
end;

procedure TZSheetPrintTitles.SetFrom(const Value: word);
begin
  if Active then RequireValid(Value, FTill);

  FFrom := Value;
end;

procedure TZSheetPrintTitles.SetTill(const Value: word);
begin
  if Active then RequireValid(FFrom, Value);

  FTill := Value;
end;

function TZSheetPrintTitles.ToString: string;
var c: char;
begin
  If Active then begin
     if FColumns then c := 'C' else c := 'R';
     Result := c + IntToStr(From + 1) + ':' + c + IntToStr(Till + 1);
  end else
     Result := '';
end;

procedure TZSheetPrintTitles.RequireValid(const AFrom, ATill: word);
begin
  if not Valid(AFrom, ATill) then
     raise Exception.Create('Invalid printable titles for the worksheet.');
end;

function TZSheetPrintTitles.Valid(const AFrom, ATill: word): boolean;
var UpperLimit: word;
begin
  Result := False;
  if AFrom > ATill then exit;
  //  if AFrom < 0 then exit; // if datatype would be changed to traditional integer, then unremark

  if FColumns
     then UpperLimit := FOwner.ColCount
     else UpperLimit := FOwner.RowCount;

  If ATill >= UpperLimit then exit;

  Result := True;
end;

//Условное форматирование
{$IFDEF ZUSE_CONDITIONAL_FORMATTING}

////::::::::::::: TZConditionalStyleItem :::::::::::::::::////

constructor TZConditionalStyleItem.Create();
begin
  Clear();
end;

procedure TZConditionalStyleItem.Clear();
begin
  FCondition := ZCFCellContentIsBetween;
  FConditionOperator := ZCFOpEqual;
  FApplyStyleID := -1;
  FValue1 := '';
  FValue2 := '';
  FBaseCellPageIndex := -1;
  FBaseCellRowIndex := 0;
  FBaseCellColumnIndex := 0;
end; //Clear

procedure TZConditionalStyleItem.Assign(Source: TPersistent);
var
  t: TZConditionalStyleItem;

begin
  if (Source is TZConditionalStyleItem) then
  begin
    t := (Source as TZConditionalStyleItem);
    Condition := t.Condition;
    ConditionOperator := t.ConditionOperator;
    Value1 := t.Value1;
    Value2 := t.Value2;
    ApplyStyleID := t.ApplyStyleID;
    BaseCellColumnIndex := t.BaseCellColumnIndex;
    BaseCellPageIndex := t.BaseCellPageIndex;
    BaseCellRowIndex := t.BaseCellRowIndex;
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalStyleItem.IsEqual(Source: TPersistent): boolean;
var
  t: TZConditionalStyleItem;

begin
  result := false;
  if (Source is TZConditionalStyleItem) then
  begin
    t := (Source as TZConditionalStyleItem);
    if (Condition  <> t.Condition) then
      exit;
    if (ConditionOperator <> t.ConditionOperator) then
      exit;
    if (ApplyStyleID <> t.ApplyStyleID) then
      exit;
    if (BaseCellColumnIndex <> t.BaseCellColumnIndex) then
      exit;
    if (BaseCellPageIndex <> t.BaseCellPageIndex) then
      exit;
    if (BaseCellRowIndex <> t.BaseCellRowIndex) then
      exit;
    if (Value1 <> t.Value1) then
      exit;
    if (Value2 <> t.Value2) then
      exit;
    result := true;
  end;
end; //IsEqual

////::::::::::::: TZConditionalStyle :::::::::::::::::////

constructor TZConditionalStyle.Create();
var
  i: integer;

begin
  FCount := 0;
  FMaxCount := 3;
  SetLength(FConditions, FMaxCount);
  for i := 0 to FMaxCount - 1 do
    FConditions[i] := TZConditionalStyleItem.Create();
  FAreas := TZConditionalAreas.Create();
end; //Create

destructor TZConditionalStyle.Destroy();
var
  i: integer;

begin
  for i := 0 to FMaxCount - 1 do
    if (Assigned(FConditions[i])) then
      FreeAndNil(FConditions[i]);
  SetLength(FConditions, 0);
  FConditions := nil;
  FreeAndNil(FAreas);
  inherited;
end; //Destroy

procedure TZConditionalStyle.Assign(Source: TPersistent);
var
  t: TZConditionalStyle;
  i: integer;

begin
  if (Source is TZConditionalStyle) then
  begin
    t := Source as TZConditionalStyle;
    Count := t.Count;
    FAreas.Assign(t.Areas);
    for i := 0 to Count - 1 do
      FConditions[i].Assign(t.Items[i]);
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalStyle.IsEqual(Source: TPersistent): boolean;
var
  t: TZConditionalStyle;
  i: integer;

begin
  result := false;
  if (Source is TZConditionalStyle) then
  begin
    t := Source as TZConditionalStyle;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FConditions[i].IsEqual(t.Items[i])) then
        exit;
    result := true;
  end;
end; //IsEqual

function TZConditionalStyle.GetItem(num: integer): TZConditionalStyleItem;
begin
  result := nil;
  if (num >= 0) and (num < Count) then
    result := FConditions[num];
end; //GetItem

procedure TZConditionalStyle.SetItem(num: integer; Value: TZConditionalStyleItem);
begin
  if (num >= 0) and (num < Count) then
    FConditions[num].Assign(Value);
end; //SetItem

procedure TZConditionalStyle.SetCount(value: integer);
var
  i: integer;

begin
  {tut}
  //TODO: нужно ли ограничение на максимальное кол-во?
  if (value >= 0) then
  begin
    if (value < FCount) then
    begin
      for i := value to FCount - 1 do
        FConditions[i].Clear();
    end else
    if (value > FMaxCount) then
    begin
      SetLength(FConditions, value);
      for i := FMaxCount to value - 1 do
        FConditions[i] := TZConditionalStyleItem.Create();
      FMaxCount := value;
    end;
    FCount := value;
  end;
end; //SetCount

procedure TZConditionalStyle.SetAreas(Value: TZConditionalAreas);
begin
  if (Assigned(Value)) then
    FAreas.Assign(Value);
end;

function TZConditionalStyle.Add(): TZConditionalStyleItem;
begin
  Count := Count + 1;
  result := FConditions[Count - 1];
end; //Add

function TZConditionalStyle.Add(StyleItem: TZConditionalStyleItem): TZConditionalStyleItem;
begin
  result := Add();
  if (Assigned(StyleItem)) then
    result.Assign(StyleItem);
end; //Add

procedure TZConditionalStyle.Delete(num: integer);
var
  i: integer;
  t: TZConditionalStyleItem;

begin
  if (num >= 0) and (num < Count) then
  begin
    t := FConditions[num];
    for i := num to Count - 2 do
      FConditions[i] := FConditions[i + 1];
    if (Count > 0) then
      FConditions[Count - 1] := t;
    Count := Count - 1;
  end;
end; //Delete

procedure TZConditionalStyle.Insert(num: integer);
begin
  Insert(num, nil);
end; //Insert

procedure TZConditionalStyle.Insert(num: integer; StyleItem: TZConditionalStyleItem);
var
  i: integer;
  t: TZConditionalStyleItem;

begin
  if (num >= 0) and (num < Count) then
  begin
    Add();
    t := FConditions[Count - 1];
    for i := Count - 1 downto num + 1 do
      FConditions[i] := FConditions[i - 1];
    FConditions[num] := t;
    if (Assigned(StyleItem)) then
      FConditions[num].Assign(StyleItem);
  end;
end; //Insert

////::::::::::::: TZConditionalAreaItem :::::::::::::::::////

constructor TZConditionalAreaItem.Create();
begin
  Create(0, 0, 1, 1);
end;

constructor TZConditionalAreaItem.Create(ColumnNum, RowNum, AreaWidth, AreaHeight: integer);
begin
  Row := RowNum;
  Column := ColumnNum;
  Width := AreaWidth;
  Height := AreaHeight;
end;

procedure TZConditionalAreaItem.SetRow(Value: integer);
begin
  if (Value >= 0) then
    FRow := Value;
end; //SetRow

procedure TZConditionalAreaItem.SetColumn(Value: integer);
begin
  if (Value >= 0) then
    FColumn := Value;
end; //SetColumn

procedure TZConditionalAreaItem.SetWidth(Value: integer);
begin
  if (Value >= 0) then
    FWidth := Value;
end; //SetWidth

procedure TZConditionalAreaItem.SetHeight(Value: integer);
begin
  if (Value >= 0) then
    FHeight := Value;
end; //SetHeight

procedure TZConditionalAreaItem.Assign(Source: TPersistent);
var
  t: TZConditionalAreaItem;

begin
  if (Source is TZConditionalAreaItem) then
  begin
    t := Source as TZConditionalAreaItem;
    Row := t.Row;
    Column := t.Column;
    Height := t.Height;
    Width := t.Width;
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalAreaItem.IsEqual(Source: TPersistent): boolean;
var
  t: TZConditionalAreaItem;

begin
  result := false;
  if (Source is TZConditionalAreaItem) then
  begin
    t := Source as TZConditionalAreaItem;
    if (FRow <> t.Row) then
      exit;
    if (FColumn <> t.Column) then
      exit;
    if (FWidth <> t.Width) then
      exit;
    if (FHeight <> t.Height) then
      exit;

    result := true;
  end;
end; //IsEqual

////::::::::::::: TZConditionalAreas :::::::::::::::::////

constructor TZConditionalAreas.Create();
begin
  FCount := 1;
  SetLength(FItems, FCount);
  FItems[0] := TZConditionalAreaItem.Create();
end; //Create

destructor TZConditionalAreas.Destroy();
var
  i: integer;

begin
  for i := 0 to FCount - 1 do
    if (Assigned(FItems[i])) then
      FreeAndNil(FItems[i]);
  Setlength(FItems, 0);
  FItems := nil;

  inherited;
end; //Destroy

procedure TZConditionalAreas.SetCount(Value: integer);
var
  i: integer;

begin
  if ((Value >= 0) and (Value <> Count)) then
  begin
    if (Value < Count) then
    begin
      for i := Value to Count - 1 do
        if (Assigned(FItems[i])) then
          FreeAndNil(FItems[i]);
      Setlength(FItems, Value);
    end else
    if (Value > Count) then
    begin
      Setlength(FItems, Value);
      for i := Count to Value - 1 do
        FItems[i] := TZConditionalAreaItem.Create();
    end;
    FCount := Value;
  end;
end; //SetCount

function TZConditionalAreas.GetItem(num: integer): TZConditionalAreaItem;
begin
  result := nil;
  if ((num >= 0) and (num < FCount)) then
    result := FItems[num];
end; //GetItem

procedure TZConditionalAreas.SetItem(num: integer; Value: TZConditionalAreaItem);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FItems[num].Assign(Value);
end; //SetItem

procedure TZConditionalAreas.Add();
begin
  SetCount(FCount + 1);
end; //Add

procedure TZConditionalAreas.Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer);
begin
  Add();
  FItems[Count - 1].Row := RowNum;
  FItems[Count - 1].Column := ColumnNum;
  FItems[Count - 1].Width := AreaWidth;
  FItems[Count - 1].Height := AreaHeight;
end; //Add

procedure TZConditionalAreas.Assign(Source: TPersistent);
var
  t: TZConditionalAreas;
  i: integer;

begin
  if (Source is TZConditionalAreas) then
  begin
    t := Source as TZConditionalAreas;
    Count := t.Count;
    for i := 0 to Count - 1 do
        FItems[i].Assign(t.Items[i]);
  end else
    inherited Assign(Source);
end; //Assign

procedure TZConditionalAreas.Delete(num: integer);
var
  t: TZConditionalAreaItem;
  i: integer;

begin
  if ((num >= 0) and (num < Count)) then
  begin
    t := FItems[num];
    for i := num to Count - 2 do
      FItems[i] := FItems[i + 1];
    FItems[Count - 1] := t;
    Count := Count -1;
  end;
end; //Delete

//Определяет, находится ли ячейка в области
//INPUT
//      ColumnNum: integer  - номер столбца ячейки
//      RowNum: integer     - номер строки ячейки
//RETURN
//      boolean - true - ячейка входит в область
function TZConditionalAreas.IsCellInArea(ColumnNum, RowNum: integer): boolean;
var
  i: integer;
  x, y, xx, yy: integer;

begin
  result := false;
  for i := 0 to Count - 1 do
  begin
    x := FItems[i].Column;
    y := FItems[i].Row;
    xx := x + FItems[i].Width;
    yy := y + FItems[i].Height;
    if ((ColumnNum >= x) and (ColumnNum < xx) and
        (RowNum >= y) and (RowNum < yy))  then
    begin
      result := true;
      break;
    end;
  end;
end; //IsCellInArea

function TZConditionalAreas.IsEqual(Source: TPersistent): boolean;
var
  t: TZConditionalAreas;
  i: integer;

begin
  result := false;
  if (Source is TZConditionalAreas) then
  begin
    t := Source as TZConditionalAreas;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FItems[i].IsEqual(t.Items[i])) then
        exit;
    result := true;
  end;
end; //IsEqual

////::::::::::::: TConditionalFormatting :::::::::::::::::////

constructor TZConditionalFormatting.Create();
begin
  FCount := 0;
  SetLength(FStyles, 0);
end; //Create

destructor TZConditionalFormatting.Destroy();
var
  i: integer;

begin
  for i := 0 to FCount - 1 do
    if (Assigned(FStyles[i])) then
      FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FStyles := nil;
  inherited;
end; //Destroy

procedure TZConditionalFormatting.SetCount(Value: integer);
var
  i: integer;

begin
  if (Value >= 0) then
  begin
    if (Value < FCount) then
    begin
      for i := Value to FCount - 1 do
         FreeAndNil(FStyles[i]);
      SetLength(FStyles, Value);
    end else
    if (Value > FCount) then
    begin
      SetLength(FStyles, Value);
      for i := FCount to Value - 1 do
        FStyles[i] := TZConditionalStyle.Create();
    end;
    FCount := value;
  end;
end; //SetCount

function TZConditionalFormatting.Add(): TZConditionalStyle;
begin
  result := Add(nil);
end; //Add

function TZConditionalFormatting.GetItem(num: integer): TZConditionalStyle;
begin
  result := nil;
  if ((num >= 0) and (num < Count)) then
    result := FStyles[num];
end; //GetItem

procedure TZConditionalFormatting.SetItem(num: integer; Value: TZConditionalStyle);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FStyles[num].Assign(Value);
end; //SetItem

function TZConditionalFormatting.Add(Style: TZConditionalStyle): TZConditionalStyle;
begin
  Count := Count + 1;
  result := FStyles[Count - 1];
  if (Assigned(Style)) then
    result.Assign(Style);
end; //Add

//Добавить условное форматирование с областью
//INPUT
//      ColumnNum: integer  - номер колонки
//      RowNum: integer     - номер строки
//      AreaWidth: integer  - ширина области
//      AreaHeight: integer - высота области
//RETURN
//      TZConditionalStyle - добавленный стиль
function TZConditionalFormatting.Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer): TZConditionalStyle;
var
  t: TZConditionalAreaItem;

begin
  result := Add(nil);
  t := result.Areas[0];
  t.Row := RowNum;
  t.Column := ColumnNum;
  t.Width := AreaWidth;
  t.Height := AreaHeight;
end; //add

procedure TZConditionalFormatting.Assign(Source: TPersistent);
var
  t: TZConditionalFormatting;
  i: integer;

begin
  if (Source is TZConditionalFormatting) then
  begin
    t := Source as TZConditionalFormatting;
    FCount := t.Count;
    for i := 0 to FCount - 1 do
      FStyles[i].Assign(t.Items[i]);
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalFormatting.IsEqual(Source: TPersistent): boolean;
var
  t: TZConditionalFormatting;
  i: integer;

begin
  result := false;
  if (Source is TZConditionalFormatting) then
  begin
    t := Source as TZConditionalFormatting;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FStyles[i].IsEqual(t.Items[i])) then
        exit;

    result := true;
  end;
end; //IsEqual

{$ENDIF} //ZUSE_CONDITIONAL_FORMATTING

{$IFDEF FPC}
initialization
  {$I zexmlss.lrs}
{$ENDIF}

{
GetDeviceCaps(hdc, HORZSIZE) / GetDeviceCaps(hdc, HORZRES); //горизонтальный размер пикселя в миллиметрах
GetDeviceCaps(hdc, VERTSIZE) / GetDeviceCaps(hdc, VERTRES); //вертикальный размер пикселя в миллиметрах
}

{
  Paper Size Table
 Index    Paper type               Paper size
 ----------------------------------------------------
 0        Undefined
 1        Letter                   8 1/2" x 11"
 2        Letter small             8 1/2" x 11"
 3        Tabloid                     11" x 17"
 4        Ledger                      17" x 11"
 5        Legal                    8 1/2" x 14"
 6        Statement                5 1/2" x 8 1/2"
 7        Executive                7 1/4" x 10 1/2"
 8        A3                        297mm x 420mm
 9        A4                        210mm x 297mm
 10       A4 small                  210mm x 297mm
 11       A5                        148mm x 210mm
 12       B4                        250mm x 354mm
 13       B5                        182mm x 257mm
 14       Folio                    8 1/2" x 13"
 15       Quarto                    215mm x 275mm
 16                                   10" x 14"
 17                                   11" x 17"
 18       Note                     8 1/2" x 11"
 19       #9 Envelope              3 7/8" x 8 7/8"
 20       #10 Envelope             4 1/8" x 9 1/2"
 21       #11 Envelope             4 1/2" x 10 3/8"
 22       #12 Envelope             4 3/4" x 11"
 23       #14 Envelope                 5" x 11 1/2"
 24       C Sheet                     17" x 22"
 25       D Sheet                     22" x 34"
 26       E Sheet                     34" x 44"
 27       DL Envelope               110mm x 220mm
 28       C5 Envelope               162mm x 229mm
 29       C3 Envelope               324mm x 458mm
 30       C4 Envelope               229mm x 324mm
 31       C6 Envelope               114mm x 162mm
 32       C65 Envelope              114mm x 229mm
 33       B4 Envelope               250mm x 353mm
 34       B5 Envelope               176mm x 250mm
 35       B6 Envelope               125mm x 176mm
 36       Italy Envelope            110mm x 230mm
 37       Monarch Envelope         3 7/8" x 7 1/2"
 38       6 3/4 Envelope           3 5/8" x 6 1/2"
 39       US Standard Fanfold     14 7/8" x 11"
 40       German Std. Fanfold      8 1/2" x 12"
 41       German Legal Fanfold     8 1/2" x 13"
}

end.


