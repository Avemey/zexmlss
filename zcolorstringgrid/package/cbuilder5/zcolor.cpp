//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USERES("zcolor.res");
USEPACKAGE("vcl50.bpi");
USEUNIT("..\..\src\zcftext.pas");
USEUNIT("..\..\src\zclab.pas");
USERES("..\..\src\zclabel.dcr");
USEUNIT("..\..\src\ZColorStringGrid.pas");
USERES("..\..\src\ZColorStringGrid.dcr");
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------

//   Package source.
//---------------------------------------------------------------------------

#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void*)
{
        return 1;
}
//---------------------------------------------------------------------------
