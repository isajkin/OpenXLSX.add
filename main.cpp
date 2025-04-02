#include "myopenxlsx.h"
int main()
{
	XLDocument1 doc;
	doc.create("./Demo.xlsx", XLForceOverwrite);
	auto wks = doc.workbook().worksheet("Sheet1");
	auto c = wks.cell("A1");
	c.value() = c.font().name();
	auto f = c.font();
	f.setName((char*)"Times New Roman");
	c = wks.cell("B1");
	c.value() = "size";
	c.font().setSize(20);
	c = wks.cell("C1");
	c.value() = "bold";
	c.font().setBold(true);
	c = wks.cell("D1");
	c.value() = "italic";
	c.font().setItalic(true);
	c = wks.cell("E1");
	c.value() = "strike";
	c.font().setStrikethrough(true);
	c = wks.cell("F1");
	c.value() = "uline";
	c.font().setUnderline(1);
	c = wks.cell("G1");
	c.value() = "uline2";
	c.font().setUnderline(2);
	c = wks.cell("H1");
	c.value() = "super";
	c.font().setSuperscript(true);
	c = wks.cell("I1");
	c.value() = "sub";
	c.font().setSubscript(true);

	c = wks.cell("J1");
	c.value() = "all";
	f = c.font();
	f.setSize(20);
	f.setBold(true);
	f.setItalic(true);
	f.setStrikethrough(true);

	c = wks.cell("A2");
	c.value() = "left";
	c.setHorizontalAlignment("left");
	c = wks.cell("B2");
	c.value() = "center";
	c.setHorizontalAlignment("center");
	c = wks.cell("C2");
	c.value() = "right";
	c.setHorizontalAlignment("right");

	c = wks.cell("D2");
	c.value() = "top";
	c.setVerticalAlignment("top");
	c = wks.cell("E2");
	c.value() = "center";
	c.setVerticalAlignment("center");
	c = wks.cell("F2");
	c.value() = "bottom";
	c.setVerticalAlignment("bottom");

	c = wks.cell("G2");
	c.value() = "wrap text wrap text";
	c.setWraptext(true);
	c = wks.cell("H2");
	c.value() = "shrink shrink shrink";
	c.setShrinktofit(true);

	c = wks.cell("B3");
	c.value() = "border left";
	c.borders(0).setLineStyle(1);

	c = wks.cell("D3");
	c.value() = "border right";
	c.borders(1).setLineStyle(2);

	c = wks.cell("F3");
	c.value() = "border top";
	c.borders(2).setLineStyle(3);

	c = wks.cell("H3");
	c.value() = "border bottom";
	c.borders(3).setLineStyle(4);

	c = wks.cell("J3");
	c.value() = "diagonalUp";
	c.borders(4).setLineStyle(5);

	c = wks.cell("L3");
	c.value() = "diagonalDowm";
	c.borders(5).setLineStyle(6);


	c=wks.cell("A5");
	c.value() = "text";
	c.setNumberFormat("@");
	c = wks.cell("B5");
	c.value() = 1.23;
	c.setNumberFormat("000");
	c = wks.cell("C5");
	c.value() = 1.23;
	c.setNumberFormat("#0.0");
	c = wks.cell("D5");
	c.value() = 1.23;
	c.setNumberFormat("#0.00");
	c = wks.cell("E5");
	c.value() = 1234567890.123;
	c.setNumberFormat("###,##0.000");
	c = wks.cell("G5");
	c.value() = "01.04.2025";
	c.setNumberFormat("yyyy.mm.dd");


	c = wks.cell("B7");
	c.value() = "size bold wrap text merge";

	auto r = wks.range("B7:C8");
	r.font().setBold(true);
	r.font().setSize(16);
	r.setWraptext(true);


	for(int32_t i=0;i<4;i++)r.borders(i).setLineStyle(1);
	wks.cell("D7").value() = r.address();
	wks.merge("B7:C8");
	//	r->setHorizontalAlignment("center");

	doc.save();
	doc.close();
	return 0;
}