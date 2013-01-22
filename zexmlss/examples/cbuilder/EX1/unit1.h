//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <zexmlss.hpp>
#include <zexmlssutils.hpp>
#include <zeodfs.hpp>
#include <zeformula.hpp>
#include <zsspxml.hpp>

#include <ZColorStringGrid.hpp>
#include <Grids.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TZEXMLSS *ZEXMLSS1;
        TButton *btnCreate;
        TButton *btnFormula;
        void __fastcall btnCreateClick(TObject *Sender);
        void __fastcall btnFormulaClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
