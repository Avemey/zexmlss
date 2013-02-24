//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "zexmlss"
#pragma link "ZColorStringGrid"
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::btnCreateClick(TObject *Sender)
{
     	TZEXMLSS *XMLSS = NULL;
	__try
	{
		TAnsiToCPConverter TextConverter = NULL;
		#if __BORLANDC__ < 0x613 // < RAD Studio 2009
		TextConverter = *AnsiToUtf8;
		#endif

		XMLSS = new TZEXMLSS(NULL);

		//� ��������� 2 ��������
		XMLSS->Sheets->Count = 2;
		XMLSS->Sheets->Sheet[0]->Title = "�������� �������";
		//������� �����
		XMLSS->Styles->Count = 13;
		//0 - ��� ��������� (20)
		XMLSS->Styles->Items[0]->Font->Size = 20;
		XMLSS->Styles->Items[0]->Font->Style =  TFontStyles() << fsBold;
		XMLSS->Styles->Items[0]->Font->Name = "Tahoma";
		XMLSS->Styles->Items[0]->BGColor = 0xCCFFCC;
		XMLSS->Styles->Items[0]->CellPattern = ZPSolid;
		XMLSS->Styles->Items[0]->Alignment->Horizontal = ZHCenter;
		XMLSS->Styles->Items[0]->Alignment->Vertical = ZVCenter;
		XMLSS->Styles->Items[0]->Alignment->WrapText = true;
		//1 - ����� ��� �������
		XMLSS->Styles->Items[1]->Border->Border[0]->Weight = 1;
		XMLSS->Styles->Items[1]->Border->Border[0]->LineStyle = ZEContinuous;
		int i = 0;
		for (i = 1; i < 4; i++)
		{
			XMLSS->Styles->Items[1]->Border->Border[i]->Assign(XMLSS->Styles->Items[1]->Border->Border[0]);
		}
		//2 - ����� ��� ��������� ������� (������ �� ������)
		XMLSS->Styles->Items[2]->Assign(XMLSS->Styles->Items[1]);
		XMLSS->Styles->Items[2]->Font->Style = TFontStyles() << fsBold;
		XMLSS->Styles->Items[2]->Alignment->Horizontal = ZHCenter;
		//3-�� �����
		XMLSS->Styles->Items[3]->Font->Size = 16;
		XMLSS->Styles->Items[3]->Font->Name = "Arial Black";
		//4-�� �����
		XMLSS->Styles->Items[4]->Font->Size = 18;
		XMLSS->Styles->Items[4]->Font->Name = "Arial";
		//5-�� �����
		XMLSS->Styles->Items[5]->Font->Size = 14;
		XMLSS->Styles->Items[5]->Font->Name = "Arial Black";
        //6-��
		XMLSS->Styles->Items[6]->Assign(XMLSS->Styles->Items[1]);
		for (i = 1; i < 4; i++)
		{
			XMLSS->Styles->Items[6]->Border->Border[i]->Weight = 3;
		}
		XMLSS->Styles->Items[6]->Border->Border[5]->Weight = 2;
		XMLSS->Styles->Items[6]->Border->Border[0]->Color = clRed;
		XMLSS->Styles->Items[6]->Border->Border[5]->LineStyle = ZEContinuous;

		XMLSS->Styles->Items[6]->Border->Border[5]->Color = clGreen;
		XMLSS->Styles->Items[6]->BGColor = clYellow;
		XMLSS->Styles->Items[6]->CellPattern = ZPSolid;
		XMLSS->Styles->Items[6]->Font->Color = clRed;

		//���������� ����� �������
		for (i = 7; i < 13; i++)
		{
			XMLSS->Styles->Items[i]->Assign(XMLSS->Styles->Items[1]);
		}

		//�������������� ������������ (�����, �����, ������);
		XMLSS->Styles->Items[7]->Alignment->Horizontal = ZHLeft;
		XMLSS->Styles->Items[8]->Alignment->Horizontal = ZHCenter;
		XMLSS->Styles->Items[9]->Alignment->Horizontal = ZHRight;

		//������������ ������������ (������, �����, �����);
		XMLSS->Styles->Items[10]->Alignment->Vertical = ZVTop;
		XMLSS->Styles->Items[11]->Alignment->Vertical = ZVCenter;
		XMLSS->Styles->Items[12]->Alignment->Vertical = ZVBottom;

		//���������� ����� � ��������
		XMLSS->Sheets->Sheet[0]->RowCount = 50;
		TZSheet * sh = XMLSS->Sheets->Sheet[0];
		sh->ColCount = 20;

		sh->Cell[0][0]->CellStyle = 3;
		sh->Cell[0][0]->Data = "������ ������������� zexmlss";

		//������ � �������
		sh->Cell[0][2]->CellStyle = 4;
		sh->Cell[0][2]->Data = "�������� ��������: http://avemey.com";
		sh->Cell[0][2]->HRef = "http://avemey.com";
		sh->Cell[0][2]->HRefScreenTip = String("�������� �� ������") + sLineBreak + String("(����� ����������� ��������� � ������)");
		//����������� ������
		sh->MergeCells->AddRectXY(0, 2, 10, 2);

		//����������
		sh->Cell[0][3]->CellStyle = 5;
		sh->Cell[0][3]->Data = "���� ����������!";
		sh->Cell[0][3]->Comment = "����� ���������� 1";
		sh->Cell[0][3]->CommentAuthor = "���-�� 1 ������� ����������";
		sh->Cell[0][3]->ShowComment = true;

		sh->Cell[10][3]->CellStyle = 5;
		sh->Cell[10][3]->Data = "���� ����������!";
		sh->Cell[10][3]->Comment = "����� ���������� 2";
		sh->Cell[10][3]->CommentAuthor = "���-�� 2 ������� ����������";
		sh->Cell[10][3]->ShowComment = true;
		sh->Cell[10][3]->AlwaysShowComment = true;
		sh->ColWidths[10] = 160; //������ �������
		sh->Columns[10]->WidthMM = 40; //~40 ��

		//����������� ������
		sh->Cell[2][1]->CellStyle = 0;
		sh->Cell[2][1]->Data = "�����-�� ���������";
		sh->MergeCells->AddRectXY(2, 1, 12, 1);

		sh->Cell[0][6]->Data = String("�����������") + sLineBreak + String("������!");
		sh->Cell[0][6]->CellStyle = 0;
		sh->MergeCells->AddRectXY(0, 6, 3, 14);

		sh->Cell[9][6]->Data = "� �����";
		sh->Cell[9][6]->CellStyle = 2;

		sh->Cell[10][6]->Data = "������ �����";
		sh->Cell[10][6]->CellStyle = 2;

		int j = 8;

		for (i = 0; i < 7; i++)
		{
		  sh->Cell[9][j]->Data = IntToStr(i) + String("-��");
		  sh->Cell[10][j]->Data = "�����";
		  sh->Cell[10][j]->CellStyle = i;
		  j+=2;
		}

		sh->Columns[5]->WidthMM = 30;
		sh->Columns[6]->WidthMM = 30;

		sh->Cell[5][5]->Data = "������������";
		sh->MergeCells->AddRectXY(5, 5, 6, 5);

		sh->Cell[5][6]->Data = "�����������";
		sh->Cell[5][6]->CellStyle = 2;

		sh->Cell[6][6]->Data = "���������";
		sh->Cell[6][6]->CellStyle = 2;

		for (i = 5; i < 7; i++)
		for (j = 7; j < 10; j++)
		  sh->Cell[i][j]->Data = "text";

		for (i = 7; i < 10; i++)
		{
		  sh->Cell[5][i]->CellStyle = i;
		  sh->Cell[6][i]->CellStyle = i + 3;
		  sh->Rows[i]->HeightMM = 14;
		}

		//����������� � 0 �� 1-�� ��������
		XMLSS->Sheets->Sheet[1]->Assign(sh);
		XMLSS->Sheets->Sheet[1]->Title = "�������� ������� (�����)";

		//��������� 0-�� � 1-�� ��������

//		SaveXmlssToEXML(XMLSS, /*�����-�� ����*/"save_test.xml", (0, 1), 2, (), 0, TextConverter, "UTF-8", "");
		int sheets[2] = {0, 1};
		String sheetnames[2] = {String("1"), String("2")};
		//���������� ������: SaveXmlssToEXML(XMLSS, "save_test.xml", sheets, 2, sheetnames, 0, NULL/*TextConverter*/, "UTF-8");
		SaveXmlssToEXML(XMLSS, "save_test.xml", sheets, 2, sheetnames, 0, TextConverter, "UTF-8");
		SaveXmlssToODFSPath(XMLSS, "C:\\1\\", sheets, 2, sheetnames, 0, TextConverter, "UTF-8");
		//SaveXmlssToEXML(Zexmlss::TZEXMLSS* &XMLSS, System::UnicodeString FileName, int const *SheetsNumbers, const int SheetsNumbers_Size, System::UnicodeString const *SheetsNames, const int SheetsNames_Size, Zsspxml::TAnsiToCPConverter TextConverter, System::UnicodeString CodePageName, System::AnsiString BOM = "")/* overload */;
		//������� �����:
		//	/home/user_name/some_path/
		//	d:\some_path
		//SaveXmlssToODFSPath(XMLSS, "C:\work\cpp\1\\", [0, 1], [], TextConverter, "UTF-8");

	}
	__finally
	{
		if (XMLSS != NULL)
		{
			XMLSS->Free();
		}

    }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::btnFormulaClick(TObject *Sender)
{
        TZEXMLSS *XMLSS = NULL;
	__try
	{
		TAnsiToCPConverter TextConverter = NULL;
		#if __BORLANDC__ < 0x613 // < RAD Studio 2009
		TextConverter = *AnsiToUtf8;
		#endif

		XMLSS = new TZEXMLSS(NULL);

		//� ��������� 2 ��������
		XMLSS->Sheets->Count = 2;
		XMLSS->Sheets->Sheet[0]->Title = "�������� �������";
		//������� �����
		XMLSS->Styles->Count = 13;
		//0 - ��� ��������� (20)
		XMLSS->Styles->Items[0]->Font->Size = 20;
		XMLSS->Styles->Items[0]->Font->Style =  TFontStyles() << fsBold;
		XMLSS->Styles->Items[0]->Font->Name = "Tahoma";
		XMLSS->Styles->Items[0]->BGColor = 0xCCFFCC;
		XMLSS->Styles->Items[0]->CellPattern = ZPSolid;
		XMLSS->Styles->Items[0]->Alignment->Horizontal = ZHCenter;
		XMLSS->Styles->Items[0]->Alignment->Vertical = ZVCenter;
		XMLSS->Styles->Items[0]->Alignment->WrapText = true;
		//1 - ����� ��� �������
		XMLSS->Styles->Items[1]->Border->Border[0]->Weight = 1;
		XMLSS->Styles->Items[1]->Border->Border[0]->LineStyle = ZEContinuous;
		int i = 0;
		for (i = 1; i < 4; i++)
		{
			XMLSS->Styles->Items[1]->Border->Border[i]->Assign(XMLSS->Styles->Items[1]->Border->Border[0]);
		}
		//2 - ����� ��� ��������� ������� (������ �� ������)
		XMLSS->Styles->Items[2]->Assign(XMLSS->Styles->Items[1]);
		XMLSS->Styles->Items[2]->Font->Style = TFontStyles() << fsBold;
		XMLSS->Styles->Items[2]->Alignment->Horizontal = ZHCenter;
		//3-�� �����
		XMLSS->Styles->Items[3]->Font->Size = 16;
		XMLSS->Styles->Items[3]->Font->Name = "Arial Black";
		//4-�� �����
		XMLSS->Styles->Items[4]->Font->Size = 18;
		XMLSS->Styles->Items[4]->Font->Name = "Arial";
		//5-�� �����
		XMLSS->Styles->Items[5]->Font->Size = 14;
		XMLSS->Styles->Items[5]->Font->Name = "Arial Black";
        	//6-��
		XMLSS->Styles->Items[6]->Assign(XMLSS->Styles->Items[1]);
		for (i = 1; i < 4; i++)
		{
			XMLSS->Styles->Items[6]->Border->Border[i]->Weight = 3;
		}
		XMLSS->Styles->Items[6]->Border->Border[5]->Weight = 2;
		XMLSS->Styles->Items[6]->Border->Border[0]->Color = clRed;
		XMLSS->Styles->Items[6]->Border->Border[5]->LineStyle = ZEContinuous;

		XMLSS->Styles->Items[6]->Border->Border[5]->Color = clGreen;
		XMLSS->Styles->Items[6]->BGColor = clYellow;
		XMLSS->Styles->Items[6]->CellPattern = ZPSolid;
		XMLSS->Styles->Items[6]->Font->Color = clRed;

		//���������� ����� �������
		for (i = 7; i < 13; i++)
		{
			XMLSS->Styles->Items[i]->Assign(XMLSS->Styles->Items[1]);
		}

		//�������������� ������������ (�����, �����, ������);
		XMLSS->Styles->Items[7]->Alignment->Horizontal = ZHLeft;
		XMLSS->Styles->Items[8]->Alignment->Horizontal = ZHCenter;
		XMLSS->Styles->Items[9]->Alignment->Horizontal = ZHRight;

		//������������ ������������ (������, �����, �����);
		XMLSS->Styles->Items[10]->Alignment->Vertical = ZVTop;
		XMLSS->Styles->Items[11]->Alignment->Vertical = ZVCenter;
		XMLSS->Styles->Items[12]->Alignment->Vertical = ZVBottom;

		//���������� ����� � ��������
		XMLSS->Sheets->Sheet[0]->RowCount = 50;
		TZSheet * sh = XMLSS->Sheets->Sheet[0];
		sh->ColCount = 20;

		sh->Cell[0][0]->CellStyle = 3;
		sh->Cell[0][0]->Data = "������ ������������� zexmlss";

		//������ � �������
		sh->Cell[0][2]->CellStyle = 4;
		sh->Cell[0][2]->Data = "�������� ��������: http://avemey.com";
		sh->Cell[0][2]->HRef = "http://avemey.com";
		sh->Cell[0][2]->HRefScreenTip = String("�������� �� ������") + sLineBreak + String("(����� ����������� ��������� � ������)");
		//����������� ������
		sh->MergeCells->AddRectXY(0, 2, 10, 2);

		//����������
		sh->Cell[0][3]->CellStyle = 5;
		sh->Cell[0][3]->Data = "���� ����������!";
		sh->Cell[0][3]->Comment = "����� ���������� 1";
		sh->Cell[0][3]->CommentAuthor = "���-�� 1 ������� ����������";
		sh->Cell[0][3]->ShowComment = true;

		sh->Cell[10][3]->CellStyle = 5;
		sh->Cell[10][3]->Data = "���� ����������!";
		sh->Cell[10][3]->Comment = "����� ���������� 2";
		sh->Cell[10][3]->CommentAuthor = "���-�� 2 ������� ����������";
		sh->Cell[10][3]->ShowComment = true;
		sh->Cell[10][3]->AlwaysShowComment = true;
		sh->ColWidths[10] = 160; //������ �������
		sh->Columns[10]->WidthMM = 40; //~40 ��

		//����������� ������
		sh->Cell[2][1]->CellStyle = 0;
		sh->Cell[2][1]->Data = "�����-�� ���������";
		sh->MergeCells->AddRectXY(2, 1, 12, 1);

		sh->Cell[0][6]->Data = String("�����������") + sLineBreak + String("������!");
		sh->Cell[0][6]->CellStyle = 0;
		sh->MergeCells->AddRectXY(0, 6, 3, 14);

		sh->Cell[9][6]->Data = "� �����";
		sh->Cell[9][6]->CellStyle = 2;

		sh->Cell[10][6]->Data = "������ �����";
		sh->Cell[10][6]->CellStyle = 2;

		int j = 8;

		for (i = 0; i < 7; i++)
		{
		  sh->Cell[9][j]->Data = IntToStr(i) + String("-��");
		  sh->Cell[10][j]->Data = "�����";
		  sh->Cell[10][j]->CellStyle = i;
		  j+=2;
		}

		sh->Columns[5]->WidthMM = 30;
		sh->Columns[6]->WidthMM = 30;

		sh->Cell[5][5]->Data = "������������";
		sh->MergeCells->AddRectXY(5, 5, 6, 5);

		sh->Cell[5][6]->Data = "�����������";
		sh->Cell[5][6]->CellStyle = 2;

		sh->Cell[6][6]->Data = "���������";
		sh->Cell[6][6]->CellStyle = 2;

		for (i = 5; i < 7; i++)
		for (j = 7; j < 10; j++)
		  sh->Cell[i][j]->Data = "text";

		for (i = 7; i < 10; i++)
		{
		  sh->Cell[5][i]->CellStyle = i;
		  sh->Cell[6][i]->CellStyle = i + 3;
		  sh->Rows[i]->HeightMM = 14;
		}

		//����������� � 0 �� 1-�� ��������
		XMLSS->Sheets->Sheet[1]->Assign(sh);
		XMLSS->Sheets->Sheet[1]->Title = "�������� ������� (�����)";

		//��������� 0-�� � 1-�� ��������
		int sheets[2] = {0, 1};
		String sheetnames[2] = {String("1"), String("2")};
                //��������� ��� excel xml
		SaveXmlssToEXML(XMLSS, "save_test.xml", sheets, 2, sheetnames, 0, TextConverter, "UTF-8");
                //���� ������ �������������!
                //��������� ��� �������������� ods
		SaveXmlssToODFSPath(XMLSS, "d:\\work\\ods_path\\", sheets, 2, sheetnames, 0, TextConverter, "UTF-8");
	}
	__finally
	{
		if (XMLSS != NULL)
		{
			XMLSS->Free();
		}

    	}

}
//---------------------------------------------------------------------------
