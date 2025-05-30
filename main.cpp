#include "myopenxlsx.h"
#include <stdio.h>
//using namespace std::literals::string_literals;
using namespace std::string_literals;
int main()
{
	XLDocument1 doc;

	doc.create("./Demo.xlsx", XLForceOverwrite);
	XLWorksheet1 wks = doc.workbook().worksheet(1);
	XLCell1 c = wks.cell("A1");
	c.value() = c.font().name();
	c.font().setColor("yellow");
	c.fill().setColor("green");
	c.fill().setPatternType(1);

	auto f = c.font();
	f.setName((char*)"Times New Roman");
	c = wks.cell("B1");
	c.value() = "size";
	c.font().setSize(20);

	c.fill().setColor("red");
	c.fill().setPatternType(1);

	c = wks.cell("C1");
	c.value() = "bold";
	c.font().setBold(true);
	c = wks.cell("D1");
	c.value() = "bold italic";
	c.font().setBold(true).setItalic();
	c = wks.cell("E1");
	c.value() = "strike";
	c.font().setStrikethrough().setColor("gold");
	c = wks.cell("F1");
	c.value() = "single";
	c.font().setUnderline("single").setColor("cyan");
	c = wks.cell("G1");
	c.value() = "double";
	c.font().setUnderline("double").setColor("silver");
	c = wks.cell("H1");
	c.value() = "super";
	c.font().setSuperscript();
	c = wks.cell("I1");
	c.value() = "sub";
	c.font().setSubscript();

	c = wks.cell("J1");
	c.value() = "all";
	f = c.font();
	f.setSize(20);
	f.setBold(true);
	f.setItalic(true);
	f.setStrikethrough(true);

	c = wks.cell("A2");
	c.value() = "left";
	c.setHorizontalAlignment("left").setVerticalAlignment("center");
	c = wks.cell("B2");
	c.value() = "center";
	c.setHorizontalAlignment("center").setVerticalAlignment("center");
	c = wks.cell("C2");
	c.value() = "right";
	c.setHorizontalAlignment("right").setVerticalAlignment("center");

	c = wks.cell("D2");
	c.value() = "top";
	c.setVerticalAlignment("top").setHorizontalAlignment("center");
	c = wks.cell("E2");
	c.value() = "center";
	c.setVerticalAlignment("center").setHorizontalAlignment("center");
	c = wks.cell("F2");
	c.value() = "bottom";
	c.setVerticalAlignment("bottom").setHorizontalAlignment("center");

	c = wks.cell("G2");
	c.value() = "wrap text wrap text";
	c.setWraptext();

	c = wks.cell("H2");
	c.value() = "shrink shrink shrink";
	c.setShrinktofit();

	c = wks.cell("B3");
	c.value() = "border left";
	auto b = c.borders(0);
	b.setLineStyle(1);
	b.setColor("blue");

	c = wks.cell("D3");
	c.value() = "border right";
	b = c.borders(1);
	b.setLineStyle(2);
	b.setColor("green");

	c = wks.cell("F3");
	c.value() = "border top";
	b = c.borders(2);
	b.setLineStyle(3);
	b.setColor("Gold");

	c = wks.cell("H3");
	c.value() = "border bottom";
	b = c.borders(3);
	b.setLineStyle(4);
	b.setColor("black");

	c = wks.cell("L3");
	c.value() = "diagonalUp";
	b = c.borders(6);
	b.setLineStyle(7);

	c = wks.cell("M3");
	c.value() = "diagonalDowm";
	b = c.borders(7);
	b.setLineStyle(8);

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
//	r.copyTo("D9");
	wks.merge("B7:C8");

	for (int32_t i = 0; i < 4; i++) {
		auto b = r.borders(i);
		b.setLineStyle(1);
		b.setColor("red");
	}

	c = wks.cell("E7");
	c.value() = "Privet";
	c.characters(2, 2).font().setItalic(true).setBold(true);

	c = wks.cell("G7");
	c.value() = "fontname";
	c.characters(5, 4).font().setUnderline("double").setSize(18).setColor("blue");

//	doc.workbook().addWorksheet("New");
//	doc.workbook().cloneSheet("Sheet1", "Clone");
//	doc.workbook().addWorksheet("Old");
//	doc.workbook().deleteSheet("Old");
	XLRECT rect;
	rect.left = 10;
	rect.top = 10;
	rect.right = 0;
	rect.bottom = 0;
	char* pic1 = (char *)"multfilm.jpg";
	char* pic2 = (char*)"243kur.png";
	char* pic3 = (char*)"open.ico";
//	char* pic = (char*)"gif.gif";

	wks.addPicture(pic1, &rect);
	rect.left = 7;
	wks.addPicture(pic2, &rect);
	rect.left = 5;
	wks.addPicture(pic3, &rect);



wks.pictures().item(2).setRotation(350);
wks.pictures().item(2).setHeight(200);
wks.pictures().item(2).fillRect();


int cnt = wks.pictures().count();
wks.cell("I11").value() = cnt;
wks.pictures().item(1).setName("Kartinka 1");

	wks.cell("I12").value() = wks.pictures().item(1).name();
	wks.cell("I13").value() = wks.pictures().item(2).name();

	rect.left = 5;
	rect.top = 15;
	rect.right = 6;
	rect.bottom = 17;
	XLSIZE size;
	size.cx = 0;
	size.cy = 0;


auto sps = wks.shapes();
auto l=sps.addLine(1,10,2,12);

auto sp=sps.addTextBox(1,5,5,120,40);
XLTextFrame1 tf=sp.textFrame();
tf.setHorizontalAlignment("center");
tf.setVerticalAlignment("justify");
XLCharacters2 ch = tf.Characters();
ch.setText("Privet Privet Privet Privet");
ch.font().setBold(true);
ch.font().setStrikethrough(true);
ch.font().setUnderline(2);
ch.font().setSize(12);
sp.setRotation(330);
sp = sps.addShape(9, 2, 2, 10, 10);

	doc.save();
	doc.close();
	return 0;
}