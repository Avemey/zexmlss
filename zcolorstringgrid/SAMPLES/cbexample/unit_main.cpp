//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "unit_main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "ZColorStringGrid"
#pragma link "zclab"
#pragma resource "*.dfm"
TfrmMain *frmMain;
//---------------------------------------------------------------------------
__fastcall TfrmMain::TfrmMain(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TfrmMain::FormCreate(TObject *Sender)
{
  ZSG->Cells[1][0]  = "Ячейка";
  ZSG->CellStyle[1][0]->Rotate = 180;
  ZSG->CellStyle[1][0]->HorizontalAlignment = 2;
  ZSG->MergeCells->AddRectXY(1, 0, 2, 0);

  ZSG->CellStyle[1][1]->Rotate = 60;
  ZSG->CellStyle[1][1]->Font->Size = 16;
  ZSG->Cells[1][1]  = "Поворот на 60 градусов";

  ZSG->Cells[0][1]  = "Поворот на 270 градусов";
  ZSG->CellStyle[0][1]->Rotate = 270;
  ZSG->CellStyle[0][1]->HorizontalAlignment = 2;
  ZSG->MergeCells->AddRectXY(0, 1, 0, 5);
}
//---------------------------------------------------------------------------

