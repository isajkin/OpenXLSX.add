#include "myopenxlsx.h"
#include <pugixml.hpp>
#include <XLUtilities.hpp>
#include <string.h>
#include "utf8.h"

const std::string ShapeNodeNameDr = "xdr:twoCellAnchor";

int nametorgb(char*);
int utf8next(utf8_int8_t* str);
utf8_int8_t* utf8substr(utf8_int8_t* str, int start, int len, int* outlen);
std::string xlrtf(std::string s, int32_t start, int32_t len, std::string rtf);


static XLNUMBERFORMATSTRUCTEMBED numberformatem[] =
{
0,"",
1,"0",
2,"0.00",
3,"#.##0",
4,"#,##0.00",
9,"0%",
10,"0.00%",
11,"0.00E+00",
12,"# ?/?",
13,"# ??/??",
14,"d/m/yyyy",
15,"d-mmm-y",
16,"d-mmm",
17,"mmm-yy",
18,"h:mm tt",
19,"h:mm:ss tt",
20,"H:mm",
21,"H:mm:ss",
22,"m/d/yyyy H:mm",
37,"#,##0;(#,##0)",
38,"#,##0;[Red](#,##0)",
39,"#,##0.00;(#,##0.00)",
40,"#,##0.00;[Red](#,##0.00)",
45,"mm:ss",
46,"[h]:mm:ss",
47,"mmss.0",
48,"##0.0E+0",
49,"@",
-1,""
};

static char* findnumberformatem(int id)
{
	for (int i = 0; true; i++) {
		if (numberformatem[i].id < 0)break;
		if (numberformatem[i].id == id)return numberformatem[i].formatcode;
	}
	return NULL;
}

static int32_t findnumberformatem(char* code)
{
	for (int i = 0; true; i++) {
		if (numberformatem[i].id < 0)break;
		if (!strcmp(numberformatem[i].formatcode, code))return numberformatem[i].id;
	}
	return -1;
}

static int XLAlignmentStyleFromString(std::string alignment)
{
	if (alignment == ""
		|| alignment == "general")       return XLAlignGeneral;
	if (alignment == "left")             return XLAlignLeft;
	if (alignment == "right")            return XLAlignRight;
	if (alignment == "center")           return XLAlignCenter;
	if (alignment == "top")              return XLAlignTop;
	if (alignment == "bottom")           return XLAlignBottom;
	if (alignment == "fill")             return XLAlignFill;
	if (alignment == "justify")          return XLAlignJustify;
	if (alignment == "centerContinuous") return XLAlignCenterContinuous;
	if (alignment == "distributed")      return XLAlignDistributed;
	return XLAlignInvalid;
}

static std::string XLAlignmentStyleToString(int alignment)
{
	switch (alignment) {
	case XLAlignGeneral: return "";
	case XLAlignLeft: return "left";
	case XLAlignRight: return "right";
	case XLAlignCenter: return "center";
	case XLAlignTop: return "top";
	case XLAlignBottom: return "bottom";
	case XLAlignFill: return "fill";
	case XLAlignJustify: return "justify";
	case XLAlignCenterContinuous: return "centerContinuous";
	case XLAlignDistributed: return "distributed";
	case XLAlignInvalid: [[fallthrough]];
	default: return "(unknown)";
	}
}

static XLUnderlineStyle XLUnderlineStyleFromString(std::string underline)
{
	if (underline == ""
		|| underline == "none")   return XLUnderlineNone;
	if (underline == "single") return XLUnderlineSingle;
	if (underline == "double") return XLUnderlineDouble;
	return XLUnderlineInvalid;
}

static std::string upper(std::string str)
{
	return std::string(_strupr(str.data()));

}

//----------------------class XLDocument1-----------------------------------
#ifdef MY_DRAWING
char* XLDocument1::shapeAttribute(int sheetXmlNo, int shapeNo, char* path)
{
	char* s, * s0; int i, att = 0; XMLNode f; std::string ss, sa;

	XLWorksheet wks = doc()->workbook().worksheet(sheetXmlNo);
	wks.drawing1();

	XLDrawing1 dr = doc()->sheetDrawing1((uint16_t)sheetXmlNo);
	XMLNode n = dr.shapeNode((uint32_t)shapeNo - 1);
	s = s0 = path;
	while (1) {
		if (*s == '/' || *s == '@' || !*s) {
			i = s - s0;
			if (i)ss = std::string((const char*)s0, (size_t)i);
			s0 = s + 1;
			if (att && i) {
				const pugi::char_t* a = (const pugi::char_t*)sa.data();
				if (n.attribute(a).empty())break;
				return (char*)n.attribute(a).as_string();
			}
			if (!att && i) {
				f = n.first_child_of_type(pugi::node_element);
				while (!f.empty()) {
					if (f.raw_name() == ss)break;
					f = f.next_sibling_of_type(pugi::node_element);
				}
				if (f.empty())break;
				n = f;
			}
			if (*s == '@')att = 1;
			if (!*s)return (char*)n.text().as_string();
		}
		s++;
	}
	return (char*)"";
}
#endif
#ifdef MY_DRAWING
void XLDocument1::setShapeAttribute(int sheetXmlNo, int shapeNo, char* path, char* attribute, char* value)
{
	char* s, * s0; int i, att = 0, val = 0; XMLNode f;

	std::string ss, sa, sv;

	XLWorksheet wks = doc()->workbook().worksheet(sheetXmlNo);
	wks.drawing1();

	XLDrawing1 dr = doc()->sheetDrawing1((uint16_t)sheetXmlNo);
	XMLNode n = dr.shapeNode((uint32_t)shapeNo - 1);
	s = s0 = path;
	while (1) {
		if (*s == '/' || *s == '@' || *s == '=' || *s == ',' || !*s) {
			i = s - s0;
			if (i)ss = std::string((const char*)s0, (size_t)i);
			s0 = s + 1;
			if (att && i) {
				if (*s == '=') {
					sa = ss;
				}
				else {
					if (val) {
						sv = ss;
					}
				}
			}
			if (att && val && i) {
				const pugi::char_t* a = (const pugi::char_t*)sa.data();
				if (n.attribute(a).empty())n.append_attribute(a) = sv.data();
				else n.attribute(a).set_value(sv.data());
				att = 0;
				val = 0;
			}
			else {
				if (!att && !val && i) {
					f = n.first_child_of_type(pugi::node_element);
					while (!f.empty()) {
						if (f.raw_name() == ss)break;
						f = f.next_sibling_of_type(pugi::node_element);
					}
					if (f.empty()) {
						f = n.append_child(ss.c_str());
					}
					if (*s != ',') {
						n = f;
					}
				}
			}
			if (!*s)break;
			if (*s == '@')att = 1;
			if (*s == '=')val = 1;
			if (*s == '/') {
				s0 = s + 1;
			}
		}
		s++;
	}
	if (attribute && *attribute) {
		const pugi::char_t* a = (const pugi::char_t*)attribute;
		if (n.attribute(a).empty())n.append_attribute(a) = value;
		else n.attribute(a).set_value(value);
	}
	else {
		if (value && *value)
			n.text() = value;
	}

	return;
}
#endif
#ifdef MY_DRAWING
int XLDocument1::appendPictures(int sheetXmlNo, void* buffer, int bufferlen, char* ext, XLRECT* rect)
{
	using namespace std::literals::string_literals;
	std::string xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	XMLNode root;
	int id = 1;
	std::string v;
	v.assign((const char*)buffer, (const size_t)bufferlen);

	while (1) {
		bool yes = false;
		for (auto& item : doc()->m_contentTypes.getContentDefItems()) {
			std::string picturesFilename = "xl/media/image" + std::to_string(id) + "." + item.ext();
			if (doc()->m_archive.hasEntry(picturesFilename)) {
				yes = true;
				break;
			}
		}
		if (!yes) {
			std::string picturesFilename = "xl/media/image" + std::to_string(id) + "." + std::string(ext);
			if (!doc()->m_archive.hasEntry(picturesFilename)) {
				doc()->m_archive.addEntry(picturesFilename, v);
				if (!doc()->m_contentTypes.ExtensionExists(ext))doc()->m_contentTypes.addDefault(ext, XLContentType::Image);
				break;
			}
		}
		id++;
	}

	std::string drawingsRelsFilename = std::string("xl/drawings/_rels/drawing") + std::to_string(sheetXmlNo) + std::string(".xml.rels");
	doc()->m_data.emplace_back(doc(), drawingsRelsFilename);
	doc()->m_drwRelationships = XLRelationships(doc()->getXmlData(drawingsRelsFilename, false), drawingsRelsFilename);
	constexpr const bool DO_NOT_THROW = true;

	std::string imgtarget = std::string("../media/image") + std::to_string(id) + "." + std::string(ext);
	XLRelationshipItem imgitem = doc()->m_drwRelationships.addRelationship(XLRelationshipType::Image, imgtarget);

	XLWorksheet wks = doc()->workbook().worksheet(sheetXmlNo);
	wks.drawing1();

	XLDrawing1 dr = doc()->sheetDrawing1(sheetXmlNo);
	dr.createShape();
	root = dr.shapeNode(dr.shapeCount() - 1);
	XMLNode from = root.append_child("xdr:from");
	from.append_child("xdr:col").text() = std::to_string(rect->left - 1).data();
	from.append_child("xdr:colOff").text() = "0";
	from.append_child("xdr:row").text() = std::to_string(rect->top - 1).data();
	from.append_child("xdr:rowOff").text() = "0";
	XMLNode to = root.append_child("xdr:to");
	to.append_child("xdr:col").text() = std::to_string(rect->right - 1).data();
	to.append_child("xdr:colOff").text() = "0";
	to.append_child("xdr:row").text() = std::to_string(rect->bottom - 1).data();
	to.append_child("xdr:rowOff").text() = "0";
	XMLNode pic = root.append_child("xdr:pic");
	XMLNode nvpicpr = pic.append_child("xdr:nvPicPr");
	XMLNode pname = nvpicpr.append_child("xdr:cNvPr");
	pname.append_attribute("id") = std::to_string(id).data();
	pname.append_attribute("name") = (std::string("Picture ") + std::to_string(id)).data();
	XMLNode cpname = nvpicpr.append_child("xdr:cNvPicPr");
	XMLNode piclocks = cpname.append_child("a:picLocks");
	//   piclocks.append_attribute("noChangeAspect") = "1";
	XMLNode bf = pic.append_child("xdr:blipFill");
	XMLNode blip = bf.append_child("a:blip");
	blip.append_attribute("xmlns:r") = xmlns.data();
	blip.append_attribute("r:embed") = imgitem.id().data();
	// clrChange о•†еѓЂзЃЃ вЉ≤ аі±п•† 
	//   XMLNode clr = blip.append_child("a:clrChange");
	//    XMLNode clrfrom=clr.append_child("a:clrFrom");
	//    XMLNode srgb = clrfrom.append_child("a:srgbClr");
	//    srgb.append_attribute("val")="FFFFFF";
	//    XMLNode clrto = clr.append_child("a:clrTo");
	//    srgb = clrfrom.append_child("a:srgbClr");
	//    srgb.append_attribute("val") = "FFFFFF";

	 //   XMLNode stretch= bf.append_child("a:stretch");
	//    stretch.append_child("a:fillRect");
	XMLNode sppr = pic.append_child("xdr:spPr");
	XMLNode xfrm = sppr.append_child("a:xfrm");
	// a:off и°°:ext о•†еѓЂзЃї вЉ≤ аі±н®К//    xfrm.append_child("a:off");
	//    xfrm.append_child("a:ext");
	XMLNode geom = sppr.append_child("a:prstGeom");
	geom.append_attribute("prst") = "rect";
	geom.append_child("a:avLst");
	root.append_child("xdr:clientData");

	return dr.shapeCount();
}
#endif

XLDocument1::XLDocument1()
{
	m_doc = new XLDocument();
	m_save = 0;
	m_begin = 0;
	m_numberformat = NULL;
	m_numberformatcount = 0;
	m_borders = NULL;
	m_bordercount = 0;
	m_fonts = NULL;
	m_fontcount = 0;
	m_fills = NULL;
	m_fillcount = 0;
	m_cellformat = NULL;
	m_cellformatcount = 0;
	m_characters = NULL;
	m_charactercount = 0;

};
XLDocument1::~XLDocument1() { delete m_doc; };

void XLDocument1::getallstyles()
{
	int i;
	if (m_begin)return;
	m_begin=1;

	XLCellFormats& cellformats = m_doc->styles().cellFormats();
	m_cellformatcount = cellformats.count();
	if (m_cellformatcount) {
		m_cellformat = (XLCELLFORMATSTRUCT*)calloc(1, m_cellformatcount * sizeof(XLCELLFORMATSTRUCT));
		for (i = 0; i < m_cellformatcount; i++) {
			XLCellFormat cf = cellformats[i];
			XLCELLFORMATSTRUCT* c = m_cellformat + i;
			c->numberformatid = cf.numberFormatId();
			c->fontindex = cf.fontIndex();
			c->fillindex = cf.fillIndex();
			c->borderindex = cf.borderIndex();
			c->xfid = cf.xfId();
			XLALIGNSTRUCT a = c->alignment;
			XLAlignment al = cf.alignment();
			a.horizontal = al.horizontal();
			a.indent = al.indent();
			a.justifylastline = al.justifyLastLine();
			a.readingorder = al.readingOrder();
			a.shrinktofit = al.shrinkToFit();
			a.textrotation = al.textRotation();
			a.vertical = al.vertical();
			a.wraptext = al.wrapText();
			c->unsave = 0;
		}
	}
	XLNumberFormats& nf = m_doc->styles().numberFormats();
	m_numberformatcount = nf.count();
	if (m_numberformatcount) {
		m_numberformat = (XLNUMBERFORMATSTRUCT*)calloc(1, m_numberformatcount * sizeof(XLNUMBERFORMATSTRUCT));
		if (!m_numberformat)return;
		for (i = 0; i < m_numberformatcount; i++) {
			XLNumberFormat f = nf[i];
			XLNUMBERFORMATSTRUCT* fs = m_numberformat + i;
			auto len = sizeof(XLNUMBERFORMATSTRUCT::formatcode) - 1;
			auto flen = f.formatCode().length();
			if (len < flen)flen = len;
			memcpy(fs->formatcode, f.formatCode().data(), flen);
			fs->id = f.numberFormatId();
			fs->unsave = 0;
		} 
	}
	XLBorders& bs = m_doc->styles().borders();
	m_bordercount = bs.count();
	if (m_bordercount) {
		m_borders= (XLBORDERSTRUCT*)calloc(1, m_bordercount * sizeof(XLBORDERSTRUCT));
		if (m_borders) {
			for (i = 0; i < m_bordercount; i++) {
				XLBorder b = bs[i];
				XLBORDERSTRUCT* border = m_borders + i;

				border->bottom.style = b.bottom().style();
				border->bottom.color.argb.alpha= b.bottom().color().rgb().alpha();
				border->bottom.color.argb.red = b.bottom().color().rgb().red();
				border->bottom.color.argb.green = b.bottom().color().rgb().green();
				border->bottom.color.argb.blue = b.bottom().color().rgb().blue();

				border->left.style = b.left().style();
				border->left.color.argb.alpha = b.left().color().rgb().alpha();
				border->left.color.argb.red = b.left().color().rgb().red();
				border->left.color.argb.green = b.left().color().rgb().green();
				border->left.color.argb.blue = b.left().color().rgb().blue();

				border->right.style = b.right().style();
				border->right.color.argb.alpha = b.right().color().rgb().alpha();
				border->right.color.argb.red = b.right().color().rgb().red();
				border->right.color.argb.green = b.right().color().rgb().green();
				border->right.color.argb.blue = b.right().color().rgb().blue();

				border->top.style = b.top().style();
				border->top.color.argb.alpha = b.top().color().rgb().alpha();
				border->top.color.argb.red = b.top().color().rgb().red();
				border->top.color.argb.green = b.top().color().rgb().green();
				border->top.color.argb.blue = b.top().color().rgb().blue();

				border->horizontal.style = b.horizontal().style();

				border->vertical.style = b.vertical().style();

				border->diagonal.style = b.diagonal().style();

				border->diagonaldown = b.diagonalDown();
				border->diagonalup = b.diagonalUp();
				border->unsave = 0;
			}
		}
	}
	XLFonts& fnts = m_doc->styles().fonts();
	m_fontcount = fnts.count();
	if (m_fontcount) {
		m_fonts = (XLFONTSTRUCT*)calloc(1, m_fontcount * sizeof(XLFONTSTRUCT));
		if (m_fonts) {
			for (i = 0; i < m_fontcount; i++) {
				XLFont f = fnts[i];
				XLFONTSTRUCT* fs = m_fonts + i;
				auto len = sizeof(XLFONTSTRUCT::name) - 1;
				auto flen = f.fontName().length();
				if (len < flen)flen = len;
				memcpy(fs->name, f.fontName().data(), flen);
				fs->charset = f.fontCharset();
				fs->family = f.fontFamily();
				fs->size = f.fontSize();
				fs->hascolor = f.hasFontColor();
				if (fs->hascolor) {
					fs->fg.color.alpha = f.fontColor().alpha();
					fs->fg.color.red = f.fontColor().red();
					fs->fg.color.green = f.fontColor().green();
					fs->fg.color.blue = f.fontColor().blue();
				}
				if (f.bold())fs->bold = 1;
				if (f.italic())fs->italic = 1;
				if (f.condense())fs->condense = 1;
				if (f.extend())fs->extend = 1;
				if (f.outline())fs->outline = 1;
				if (f.shadow())fs->shadow = 1;
				if (f.strikethrough())fs->strikethrough = 1;
				fs->underline = f.underline();
				fs->scheme = f.scheme();
				fs->vertalign = f.vertAlign();
				fs->unsave = 0;
			}
		}
	}
	XLFills & fls = m_doc->styles().fills();
	m_fillcount = fls.count();
	if (m_fillcount) {
		m_fills = (XLFILLSTRUCT*)calloc(1, m_fillcount * sizeof(XLFILLSTRUCT));
		if (m_fills) {
			for (i = 0; i < m_fillcount; i++) {
				XLFill f = fls[i];
				XLFILLSTRUCT* fs = m_fills + i;
				fs->filltype = f.fillType();
				switch (fs->filltype) {
				case XLGradientFill:
					fs->gradienttype = f.gradientType();
					fs->bottom = f.bottom();
					fs->degree = f.degree();
					fs->left = f.left();
					fs->right = f.right();
					fs->top = f.top();
					break;
				case XLPatternFill:
					fs->patterntype = f.patternType();
					fs->hasbgcolor = f.hasBackgroundColor();
					if (fs->hasbgcolor) {
						fs->bg.color.alpha = f.backgroundColor().alpha();
						fs->bg.color.blue = f.backgroundColor().blue();
						fs->bg.color.green = f.backgroundColor().green();
						fs->bg.color.red = f.backgroundColor().red();
					}
					fs->hasfgcolor = f.hasColor();
					if (fs->hasfgcolor) {
						fs->fg.color.alpha = f.color().alpha();
						fs->fg.color.blue = f.color().blue();
						fs->fg.color.green = f.color().green();
						fs->fg.color.red = f.color().red();
					}
					break;
				}
				fs->unsave = 0;
			}
		}
	}
}

void XLDocument1::setcharacters()
{
	for (int i = 0; i < m_charactercount; i++) {
		XLCHARACTERSTRUCT* cs = m_characters + i;
		int indexf = cs->indexf;
		if (indexf < 0)continue;
		cs->indexf = -1;
		XLCELLFORMATSTRUCT* cf = m_cellformat + indexf;
		indexf = cf->fontindex;
		XLFONTSTRUCT* f = m_fonts + indexf;

		std::string rtf = "";
		if (f->bold)rtf = rtf + "<b/>";
		if (f->italic)rtf = rtf + "<i/>";
		if (f->underline) {
			if (f->underline == 1) {
				rtf = rtf + "<u/>";
			}
			else {
				if (f->underline == 2) {
					rtf = rtf + "<u val=\"double\"/>";
				}
			}
		}
		if (f->strikethrough)rtf = rtf + "<strike/>";
		if (f->vertalign == 1)rtf = rtf + "<vertAlign val=\"subscript\"/>";
		if (f->vertalign == 2)rtf = rtf + "<vertAlign val=\"superscript\"/>";
		if (f->hascolor) {
			XLColor color(f->fg.color.alpha, f->fg.color.red, f->fg.color.green, f->fg.color.blue);
			color.hex();
			rtf = rtf + "<color rgb=\"" + color.hex() + "\"/>";
		}
		if (f->size) {
			char buf[32];
			char* s = _itoa(f->size, buf, 10);
			rtf = rtf + "<sz val=\"" + std::string(s) + "\"/>";
		}
		if (f->charset) {
			char buf[32];
			itoa(f->charset, buf, 10);
			rtf = rtf + "<charset val=\"" + std::string(buf) + "\"/>";
		}
		if (f->name[0])rtf = rtf + "<rFont val=\"" + std::string(f->name) + "\"/>";

		if (rtf.length()) {
			XLCell cell = m_doc->workbook().worksheet(cs->sheetno).cell(cs->row, cs->col);
			std::string v = cell.value().getString();
			v = xlrtf(v, cs->start, cs->len, rtf);
			cell.value() = v;
		}
	}
}

static XLColor fromargb(XLCOLORSTRUCT *argb)
{
	XLColor c;
	memcpy(&c,argb, sizeof(XLColor));
	return c;
}

void XLDocument1::setallstyles()
{
	int i;
	if (m_save) {
		m_save=0;

		XLNumberFormats& nf = m_doc->styles().numberFormats();
		while (nf.count() < (size_t)m_numberformatcount)nf.create();
		for (i = 0; i < m_numberformatcount; i++) {
			if (!m_numberformat[i].unsave)continue;
			nf[i].setNumberFormatId(m_numberformat[i].id);
			nf[i].setFormatCode(m_numberformat[i].formatcode);
			m_numberformat[i].unsave = 0;
		}
		XLBorders bs = m_doc->styles().borders();
		while (bs.count() < (size_t)m_bordercount) {
			bs.create();
		}
		for (i = 0; i < m_bordercount; i++) {
			XLBORDERSTRUCT* border = m_borders + i;
			if (!border->unsave)continue;
			XLBorder b = bs.borderByIndex(i);
			if(border->bottom.style)b.setBottom((XLLineStyle)border->bottom.style, fromargb(&border->bottom.color.argb), 0);
			if(border->left.style)b.setLeft((XLLineStyle)border->left.style, fromargb(&border->left.color.argb), 0);
			if(border->right.style)b.setRight((XLLineStyle)border->right.style, fromargb(&border->right.color.argb), 0);
			if(border->top.style)b.setTop((XLLineStyle)border->top.style, fromargb(&border->top.color.argb), 0);
			if (border->horizontal.style)b.setHorizontal((XLLineStyle)border->horizontal.style, fromargb(&border->horizontal.color.argb), 0);
			if (border->vertical.style)b.setVertical((XLLineStyle)border->vertical.style, fromargb(&border->vertical.color.argb), 0);
			if (border->diagonal.style)b.setDiagonal((XLLineStyle)border->diagonal.style, fromargb(&border->diagonal.color.argb), 0);
			if(border->diagonaldown)b.setDiagonalDown(border->diagonaldown);
			if(border->diagonalup)b.setDiagonalUp(border->diagonalup);
			border->unsave = 0;
		}
		XLFonts& fnts = m_doc->styles().fonts();
		while (fnts.count() < (size_t)m_fontcount)fnts.create();
		for (i = 0; i < m_fontcount; i++) {
			XLFONTSTRUCT *f = m_fonts+i; XLFont fs = fnts[i];
			if (!f->unsave)continue;
			if (f->bold)fs.setBold(f->bold);
			if (f->italic)fs.setItalic(f->italic);
			if (f->name[0])fs.setFontName(f->name);
			if (f->size)fs.setFontSize(f->size);
			if (f->charset)fs.setFontCharset(f->charset);
			if (f->family)fs.setFontFamily(f->family);
			if (f->hascolor) {
				XLColor c(f->fg.color.alpha, f->fg.color.red, f->fg.color.green, f->fg.color.blue);
				fs.setFontColor(c);
			}
			if (f->condense)fs.setCondense(f->condense);
			if (f->extend)fs.setExtend(f->extend);
			if (f->outline)fs.setOutline(f->outline);
			if (f->shadow)fs.setShadow(f->shadow);
			if (f->strikethrough)fs.setStrikethrough(f->strikethrough);
			if (f->underline)fs.setUnderline((XLUnderlineStyle)f->underline);
			if (f->scheme)fs.setScheme((XLFontSchemeStyle)f->scheme);
			if (f->vertalign)fs.setVertAlign((XLVerticalAlignRunStyle)f->vertalign);
			f->unsave = 0;
		}
		XLFills& fls = m_doc->styles().fills();
		while (fls.count() < (size_t)m_fillcount)fls.create();
		for (i = 0; i < m_fillcount; i++) {
			XLFILLSTRUCT* f = m_fills + i; XLFill fs = fls[i];
			if (!f->unsave)continue;
			fs.setFillType((XLFillType)f->filltype);
			switch (f->filltype) {
			case XLGradientFill:
				fs.setGradientType((XLGradientType)f->gradienttype);
				fs.setBottom(f->bottom);
				fs.setDegree(f->degree);
				fs.setLeft(f->left);
				fs.setRight(f->right);
				fs.setTop(f->top);
				break;
			case XLPatternFill:
				fs.setPatternType((XLPatternType)f->patterntype);
				if (f->hasfgcolor) {
					XLColor c(f->fg.color.alpha, f->fg.color.red, f->fg.color.green, f->fg.color.blue);
					fs.setColor(c);
				}
				if (f->hasbgcolor) {
					XLColor c(f->bg.color.alpha, f->bg.color.red, f->bg.color.green, f->bg.color.blue);
					fs.setColor(c);
				}
				break;
			}
			f->unsave = 0;
		}

		XLCellFormats cf = m_doc->styles().cellFormats();
		while (cf.count() < (size_t)m_cellformatcount)cf.create();
		for (i = 0; i < m_cellformatcount; i++) {
			XLCellFormat c = cf[i];
			XLCELLFORMATSTRUCT* ce = m_cellformat + i;
			if (!ce->unsave)continue;
			c.setNumberFormatId(ce->numberformatid);
			if (c.numberFormatId())c.setApplyNumberFormat(true);

			c.setFontIndex(ce->fontindex);
			if (c.fontIndex())c.setApplyFont(true);

			c.setFillIndex(ce->fillindex);
			if (c.fillIndex())c.setApplyFill(true);

			c.setBorderIndex(ce->borderindex);
			if (c.borderIndex())c.setApplyBorder(true);

			XLALIGNSTRUCT* a = &ce->alignment;
			bool flag = false;
			if (a->horizontal) {
				flag = true; c.alignment(XLCreateIfMissing).setHorizontal((XLAlignmentStyle)a->horizontal);
			}
			if (a->vertical) {
				flag = true; c.alignment(XLCreateIfMissing).setVertical((XLAlignmentStyle)a->vertical);
			}
			if (a->indent) {
				flag = true; c.alignment(XLCreateIfMissing).setIndent((XLAlignmentStyle)a->indent);
			}
			if (a->justifylastline) {
				flag = true; c.alignment(XLCreateIfMissing).setJustifyLastLine((XLAlignmentStyle)a->justifylastline);
			}
			if (a->readingorder) {
				flag = true; c.alignment(XLCreateIfMissing).setReadingOrder((XLAlignmentStyle)a->readingorder);
			}
			if (a->shrinktofit) {
				flag = true; c.alignment(XLCreateIfMissing).setShrinkToFit((XLAlignmentStyle)a->shrinktofit);
			}
			if (a->textrotation) {
				flag = true; c.alignment(XLCreateIfMissing).setTextRotation((XLAlignmentStyle)a->textrotation);
			}
			if (a->wraptext) {
				flag = true; c.alignment(XLCreateIfMissing).setWrapText((XLAlignmentStyle)a->wraptext);
			}

			if (flag)c.setApplyAlignment(true);

			ce->unsave = 0;
		}
		setcharacters();
	}
	if (m_numberformat) {
		free(m_numberformat);
		m_numberformat = NULL;
		m_numberformatcount = 0;
	}

	if (m_borders) {
		free(m_borders);
		m_borders = NULL;
		m_bordercount = 0;
	}

	if (m_fonts) {
		free(m_fonts);
		m_fonts = NULL;
		m_fontcount = 0;
	}

	if (m_fills) {
		free(m_fills);
		m_fills = NULL;
		m_fillcount = 0;
	}

	if (m_cellformat) {
		free(m_cellformat);
		m_cellformat = NULL;
		m_cellformatcount = 0;
	}

	if (m_characters) {
		free(m_characters);
		m_characters = NULL;
		m_charactercount = 0;
	}

	m_begin=0;
}

bool XLDocument1::getboolstyle(int32_t index, int32_t type,int32_t prop)
{
	if (index < 0 or index >= m_cellformatcount)return false;
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID :index = m_cellformat[index].numberformatid; break;
	case MY_XLCELLFORMAT_ALIGNMENT: {
		XLALIGNSTRUCT al = m_cellformat[index].alignment;
		switch (prop) {
		case MY_ALIGN_SHRINKTOFIT:return al.shrinktofit;
		case MY_ALIGN_WRAPTEXT:return al.wraptext;
		}
		return 0;
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		int indexf = m_cellformat[index].fontindex;
		XLFONTSTRUCT f = m_fonts[indexf];
		switch (prop) {
			case MY_XLFONT_BOLD:return f.bold;
			case MY_XLFONT_ITALIC:return f.italic;
			case MY_XLFONT_CONDENSE:return f.condense;
			case MY_XLFONT_EXTEND:return f.extend;
			case MY_XLFONT_OUTLINE:return f.outline;
			case MY_XLFONT_SHADOW:return f.shadow;
			case MY_XLFONT_STRIKETHROUGH:return f.strikethrough;
			default:return false;
		}
	}
	case MY_XLCELLFORMAT_FILLINDEX:break;
	case MY_XLCELLFORMAT_BORDERINDEX: {
		int indexf = m_cellformat[index].borderindex;
		switch (prop) {
		case MY_BORDER_DIAGONALDOWN:return m_borders[index].diagonaldown;
		case MY_BORDER_DIAGONALUP:return m_borders[index].diagonalup;
		default: return false;
		}
	}
	case MY_XLCELLFORMAT_XFID:return false;
	

	}
	return false;
}

int32_t XLDocument1::getintstyle(int32_t index, int32_t type, int32_t prop)
{
	if (index < 0 or index >= m_cellformatcount)return 0;
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID: {
		int indexf;
		indexf = m_cellformat[index].numberformatid;
		switch (prop) {
		case MY_NUMBERFORMAT_ID:
			return indexf;
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		int indexf = m_cellformat[index].fontindex;
		XLFONTSTRUCT* f = m_fonts+indexf;
		switch (prop) {
		case MY_XLFONT_CHARSET:return f->charset;
		case MY_XLFONT_FAMILY:return f->family;
		case MY_XLFONT_SIZE:return f->size;
		case MY_XLFONT_UNDERLINE:return f->underline;
		case MY_XLFONT_SCHEME:return f->scheme;
		case MY_XLFONT_VERTALIGN:return f->vertalign;
		case MY_XLFONT_COLOR:return f->fg.argb;
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_FILLINDEX: {
		int indexf = m_cellformat[index].fillindex;
		XLFILLSTRUCT* f = m_fills + indexf;
		switch (prop) {
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_ALIGNMENT: {
		XLALIGNSTRUCT al = m_cellformat[index].alignment;
		switch (prop) {
		case MY_ALIGN_HORIZONTAL:return al.horizontal;
		case MY_ALIGN_VERTICAL:return al.vertical;
		case MY_ALIGN_INDENT:return al.indent;
		case MY_ALIGN_JUSTIFYLASTLINE:return al.justifylastline;
		case MY_ALIGN_READINGORDER:return al.readingorder;
		case MY_ALIGN_RELATIVEINDENT:return al.relativeindent;
		case MY_ALIGN_TEXTROTATION:return al.textrotation;
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_BORDERINDEX: {
		int indexf = m_cellformat[index].borderindex;
		switch (prop) {
			case MY_BORDER_BOTTOM:return m_borders[indexf].bottom.style;
			case MY_BORDER_LEFT:return m_borders[indexf].left.style;
			case MY_BORDER_RIGHT:return m_borders[indexf].right.style;
			case MY_BORDER_TOP:return m_borders[indexf].top.style;
			case MY_BORDER_HORIZONTAL:return m_borders[indexf].horizontal.style;
			case MY_BORDER_VERTICAL:return m_borders[indexf].vertical.style;
			case MY_BORDER_DIAGONAL:return m_borders[indexf].diagonal.style;
			default:return 0;
		}
	}
	case MY_XLCELLFORMAT_XFID:return 0;

	}
	return 0;
}

double XLDocument1::getdoublestyle(int32_t index, int32_t type, int32_t prop)
{
	if (index < 0 or index >= m_cellformatcount)return 0;
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID: {
		int indexf;
		indexf = m_cellformat[index].numberformatid;
		switch (prop) {
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		int indexf = m_cellformat[index].fontindex;
		XLFONTSTRUCT* f = m_fonts + indexf;
		switch (prop) {
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_FILLINDEX: {
		int indexf = m_cellformat[index].fillindex;
		XLFILLSTRUCT* f = m_fills + indexf;
		switch (prop) {
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_ALIGNMENT: {
		XLALIGNSTRUCT al = m_cellformat[index].alignment;
		switch (prop) {
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_BORDERINDEX: {
		int indexf = m_cellformat[index].borderindex;
		switch (prop) {
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_XFID:return 0;

	}
	return 0;
}

char* XLDocument1::findnumberformat(int id)
{
	for (int i = 0; i < m_numberformatcount; i++) {
		if (m_numberformat[i].id == id)return m_numberformat[i].formatcode;
	}
	return NULL;
}

char * XLDocument1::getcharstyle(int32_t index, int32_t type, int32_t prop)
{
	if (index < 0 or index >= m_cellformatcount)return (char *)"";
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID: {
		int indexf = m_cellformat[index].numberformatid; char* ret;
		switch (prop) {
		case MY_NUMBERFORMAT_CODE:
			ret = findnumberformatem(indexf);
			if (ret)return ret;
			ret= findnumberformat(indexf);
			if (ret)return ret;
			return (char *)"";
		}
		break;
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		int indexf = m_cellformat[index].fontindex;
		XLFONTSTRUCT* f = m_fonts + indexf;
		switch (prop) {
		case MY_XLFONT_NAME:return f->name;
		default:return (char*)"";
		}
	}
	case MY_XLCELLFORMAT_FILLINDEX:
	case MY_XLCELLFORMAT_ALIGNMENT:
	case MY_XLCELLFORMAT_BORDERINDEX:
	case MY_XLCELLFORMAT_XFID:return (char *)"";

	}
	return (char *)"";
}

int32_t XLDocument1::findnumberformat(char* code)
{
	for (int i = 0; i < m_numberformatcount; i++) {
		if (!strcmp(m_numberformat[i].formatcode, code))return m_numberformat[i].id;
	}
	return -1;
}

int32_t XLDocument1::getnumberformatnextfreeid()
{
	int32_t next = m_numberformatnextfreeid;
	for (int i = 0; i < m_numberformatcount; i++) {
		if (m_numberformat[i].id >= next) {
			m_numberformatnextfreeid = next + 1;
			next++;
		}
	}
	return next;
}

int32_t XLDocument1::createnumberformat(char* code) {
	int len = strlen(code);
	if (!len || len > sizeof(XLNUMBERFORMATSTRUCT::formatcode) - 1)return 0;
	int next = getnumberformatnextfreeid();
	m_numberformat = (XLNUMBERFORMATSTRUCT*)realloc((void*)m_numberformat, sizeof(XLNUMBERFORMATSTRUCT) * (m_numberformatcount + 1));
	if (m_numberformat) {
		strcpy(m_numberformat[m_numberformatcount].formatcode, code);
		m_numberformat[m_numberformatcount].id = next;
		m_numberformat[m_numberformatcount].unsave = 1;
		m_numberformatcount++;
		return next;
	}
	return 0;
}

int32_t XLDocument1::findfont(void* p)
{
	int i;
	for (i = 0; i < m_fontcount; i++) {
		if (!memcmp((void *)&m_fonts[i], (void *)p, sizeof(XLFONTSTRUCT)))return i;
	}
	return -1;
}

int32_t XLDocument1::createfont(void *p) {
	m_fonts = (XLFONTSTRUCT *)realloc((void*)m_fonts, sizeof(XLFONTSTRUCT) * (m_fontcount + 1));
	if (m_fonts && m_fontcount) {
		memcpy((void*)&m_fonts[m_fontcount], p, sizeof(XLFONTSTRUCT));
		m_fonts[m_fontcount].unsave = 1;
		m_fontcount++;
		return m_fontcount - 1;
	}
	return 0;
}

int32_t XLDocument1::findfill(void* p)
{
	int i;
	for (i = 0; i < m_fillcount; i++) {
		if (!memcmp((void*)&m_fills[i], (void*)p, sizeof(XLFILLSTRUCT)))return i;
	}
	return -1;
}

int32_t XLDocument1::createfill(void* p) {
	m_fills = (XLFILLSTRUCT*)realloc((void*)m_fills, sizeof(XLFILLSTRUCT) * (m_fillcount + 1));
	if (m_fills && m_fillcount) {
		memcpy((void*)&m_fills[m_fillcount], p, sizeof(XLFILLSTRUCT));
		m_fills[m_fillcount].unsave = 1;
		m_fillcount++;
		return m_fillcount - 1;
	}
	return 0;
}

int32_t XLDocument1::findcharacter(void* p)
{
	for (int i = 0; i < m_charactercount; i++) {
		if (!memcmp((void*)&m_characters[i], (void*)p, sizeof(XLCHARACTERSTRUCT)-sizeof(int32_t)))return i;
	}
	return -1;
}
int32_t XLDocument1::findcharacter(int16_t sheetno,int32_t row,int16_t col)
{
	for (int i = 0; i < m_charactercount; i++) {
		if (m_characters[i].sheetno==sheetno && m_characters[i].row==row && m_characters[i].col==col)return i;
	}
	return -1;
}

int32_t XLDocument1::createcharacter(void* p) {
	int32_t index=-1;
	for (int i = 0; i < m_charactercount; i++) {
		if (m_characters[i].indexf == -1) {
			index = i;
			break;
		}
	}
	if (index == -1) {
		m_characters = (XLCHARACTERSTRUCT*)realloc((void*)m_characters, sizeof(XLCHARACTERSTRUCT) * (m_charactercount + 1));
		if (m_characters) {
			memcpy((void*)&m_characters[m_charactercount], p, sizeof(XLCHARACTERSTRUCT));
			m_charactercount++;
			return m_charactercount - 1;
		}
	}
	else {
		memcpy((void*)&m_characters[index], p, sizeof(XLCHARACTERSTRUCT));
		return index;
	}
	return -1;
}

int32_t XLDocument1::copycharacter(int32_t index, int16_t sheetno, int32_t row, int16_t col)
{
	XLCHARACTERSTRUCT p;
	if (index < 0)return -1;
	if (index >= m_charactercount)return -1;
	memcpy(&p, &m_characters[index], sizeof(XLCHARACTERSTRUCT));
	p.sheetno = sheetno;
	p.row = row;
	p.col = col;
	return createcharacter((void*)&p);
}

int32_t XLDocument1::findborder(void* p)
{
	int i;
	for (i = 0; i < m_bordercount; i++) {
		if (!memcmp((void*)&m_borders[i], (void*)p, sizeof(XLBORDERSTRUCT)))return i;
	}
	return -1;
}

int32_t XLDocument1::createborder(void* p) {
	m_borders = (XLBORDERSTRUCT*)realloc((void*)m_borders, sizeof(XLBORDERSTRUCT) * (m_bordercount + 1));
	if (m_borders) {
		memcpy((void*)&m_borders[m_bordercount], p, sizeof(XLBORDERSTRUCT));
		m_borders[m_bordercount].unsave = 1;
		m_bordercount++;
		return m_bordercount - 1;
	}
	return 0;
}

int32_t XLDocument1::findcellformat(XLCELLFORMATSTRUCT *p)
{
	int i;
	for (i = 0; i < m_cellformatcount; i++) {
		if (!memcmp((void*)&m_cellformat[i], (void*)p, sizeof(XLCELLFORMATSTRUCT)))return i;
	}
	return -1;
}

int32_t XLDocument1::countcellformat(int32_t type,int32_t n)
{
	int i; int count = 0;
	for (i = 0; i < m_cellformatcount; i++) {
		switch (type) {
		case MY_XLCELLFORMAT_NUMBERFORMATID:if (m_cellformat[i].numberformatid == n)count++; break;
		case MY_XLCELLFORMAT_FONTINDEX:if (m_cellformat[i].fontindex == n)count++; break;
		case MY_XLCELLFORMAT_FILLINDEX:if (m_cellformat[i].fillindex == n)count++; break;
		case MY_XLCELLFORMAT_BORDERINDEX:if (m_cellformat[i].borderindex == n)count++; break;
		case MY_XLCELLFORMAT_XFID:if (m_cellformat[i].xfid == n)count++; break;
		}
	}
	return count;
}

int32_t XLDocument1::createcellformat(void* p) {
	m_cellformat = (XLCELLFORMATSTRUCT*)realloc((void*)m_cellformat, sizeof(XLCELLFORMATSTRUCT) * (m_cellformatcount + 1));
	if (m_cellformatcount) {
		memcpy((void*)&m_cellformat[m_cellformatcount], p, sizeof(XLCELLFORMATSTRUCT));
		m_cellformat[m_cellformatcount].unsave = 1;
		m_cellformatcount++;
		return m_cellformatcount - 1;
	}
	return 0;
}

int32_t XLDocument1::setboolstyle(int32_t index, int32_t type, int32_t prop, bool value)
{
	switch (type) {
		case MY_XLCELLFORMAT_NUMBERFORMATID:break;
		case MY_XLCELLFORMAT_ALIGNMENT: {
			XLCELLFORMATSTRUCT pp; XLALIGNSTRUCT* al;
			if (!index) {
				memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
				index = createcellformat((void*)&pp);
			}
			al = &m_cellformat[index].alignment;
			switch (prop) {
				case MY_ALIGN_SHRINKTOFIT:al->shrinktofit = value; break;
				case MY_ALIGN_WRAPTEXT:al->wraptext = value; break;
			}
			m_save = 1;
			return index;
		}
		case MY_XLCELLFORMAT_FONTINDEX: {
			XLFONTSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
			if (m_cellformatcount)
				indexf = m_cellformat[index].fontindex;
			else
				indexf = 0;
			if (indexf && indexf < m_fontcount)
				memcpy((void*)&p, (void*)&m_fonts[indexf], sizeof(XLFONTSTRUCT));
			else
				memset((void*)&p,0, sizeof(XLFONTSTRUCT));
			switch (prop) {
				case MY_XLFONT_BOLD:p.bold = value; break;
				case MY_XLFONT_ITALIC:p.italic = value; break;
				case MY_XLFONT_OUTLINE:p.outline = value; break;
				case MY_XLFONT_SHADOW:p.shadow = value; break;
				case MY_XLFONT_STRIKETHROUGH:p.strikethrough = value; break;
				case MY_XLFONT_EXTEND:p.extend = value; break;
				case MY_XLFONT_CONDENSE:p.condense = value; break;
				default:break;
			}
			if (index)
				c = countcellformat(type, indexf);
			else
				c = 2;
			indexf = findfont((void *)&p);
			if (indexf < 0)indexf = createfont((void*)&p);
			if (c > 1) {//если несколько €чеек используют один индекс фонта, 
				//или нулевой индекс на €чейке (первые записи в словар€х не трогаем!!!),
				// создаем новый индекс
				memcpy((void *)&pp,&m_cellformat[index],sizeof(XLCELLFORMATSTRUCT));
				pp.fontindex = indexf;
				index = createcellformat((void*)&pp);
			}
			else {//иначе просто мен€ем индекс фонта на текущей €чейке
				m_cellformat[index].fontindex = indexf;
				m_cellformat[index].unsave = 1;
			}
			m_save=1;
			return index;
		}
		case MY_XLCELLFORMAT_FILLINDEX: {
			XLFILLSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
			if (m_cellformatcount)
				indexf = m_cellformat[index].fillindex;
			else
				indexf = 0;
			if (indexf && indexf < m_fillcount)
				memcpy((void*)&p, (void*)&m_fills[indexf], sizeof(XLFILLSTRUCT));
			else
				memset((void*)&p, 0, sizeof(XLFILLSTRUCT));
			switch (prop) {
			default:break;
			}
			if (index)
				c = countcellformat(type, indexf);
			else
				c = 2;
			indexf = findfill((void*)&p);
			if (indexf < 0)indexf = createfill((void*)&p);
			if (c > 1) {//если несколько €чеек используют один индекс фонта, 
				//или нулевой индекс на €чейке (первые записи в словар€х не трогаем!!!),
				// создаем новый индекс
				memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
				pp.fillindex = indexf;
				index = createcellformat((void*)&pp);
			}
			else {//иначе просто мен€ем индекс фонта на текущей €чейке
				m_cellformat[index].fillindex = indexf;
				m_cellformat[index].unsave = 1;
			}
			m_save = 1;
			return index;
		}
		case MY_XLCELLFORMAT_BORDERINDEX:{
			XLBORDERSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
			if (m_cellformatcount)
				indexf = m_cellformat[index].borderindex;
			else
				indexf = 0;
			if (indexf && indexf < m_bordercount)
				memcpy((void*)&p, (void*)&m_borders[indexf], sizeof(XLBORDERSTRUCT));
			else
				memset((void*)&p, 0, sizeof(XLBORDERSTRUCT));
			switch (prop) {
			case MY_BORDER_DIAGONALDOWN:p.diagonaldown = value; break;
			case MY_BORDER_DIAGONALUP:p.diagonalup = value; break;
			default:break;
			}
			if (index)
				c = countcellformat(type, indexf);
			else
				c = 2;
			indexf = findborder((void*)&p);
			if (indexf < 0)indexf = createborder((void*)&p);
			if (c > 1) {//если несколько €чеек используют один индекс фонта, 
				//или нулевой индекс на €чейке (первые записи в словар€х не трогаем!!!),
				// создаем новый индекс
				memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
				pp.borderindex = indexf;
				index = createcellformat((void*)&pp);
			}
			else {//иначе просто мен€ем индекс фонта на текущей €чейке
				m_cellformat[index].borderindex = indexf;
				m_cellformat[index].unsave = 1;
			}
			m_save = 1;
			return index;
	}

		default:break;
		}
	return 0;
}

int32_t XLDocument1::setintstyle(int32_t index, int32_t type, int32_t prop, int value)
{
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID: {
		XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].numberformatid;
		else
			indexf = 0;
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.numberformatid = value;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].numberformatid = value;
			m_cellformat[index].unsave = 1;
		}
		m_save=1;
		return index;
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		XLFONTSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].fontindex;
		else
			indexf = 0;
		if (indexf>0 && indexf < m_fontcount)
			memcpy((void*)&p, (void*)&m_fonts[indexf], sizeof(XLFONTSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLFONTSTRUCT));
		switch (prop) {
		case MY_XLFONT_CHARSET:p.charset = value; break;
		case MY_XLFONT_FAMILY:p.family = value; break;
		case MY_XLFONT_SIZE:p.size = value; break;
		case MY_XLFONT_UNDERLINE:p.underline = value; break;
		case MY_XLFONT_SCHEME:p.underline = value; break;
		case MY_XLFONT_VERTALIGN:p.vertalign = value; break;
		case MY_XLFONT_COLOR:p.fg.argb = value; p.hascolor = 1; break;
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findfont((void*)&p);
		if (indexf < 0)indexf = createfont((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.fontindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].fontindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save=1;
		return index;
	}
	case MY_XLCELLFORMAT_FILLINDEX: {
		XLFILLSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].fillindex;
		else
			indexf = 0;
		if (indexf > 0 && indexf < m_fillcount)
			memcpy((void*)&p, (void*)&m_fills[indexf], sizeof(XLFILLSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLFILLSTRUCT));
		switch (prop) {
		case MY_FILL_COLOR:p.fg.argb = value; p.filltype = XLPatternFill; p.hasfgcolor = 1; break;
		case MY_FILL_BACKGROUNDCOLOR:p.bg.argb = value; p.filltype = XLPatternFill; p.hasbgcolor = 1; break;
		case MY_FILL_PATTERNTYPE:p.patterntype = value; p.filltype = XLPatternFill; break;
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findfill((void*)&p);
		if (indexf < 0)indexf = createfill((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.fillindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].fillindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_ALIGNMENT: {
		XLCELLFORMATSTRUCT pp; XLALIGNSTRUCT *al;
		if (!index) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			index = createcellformat((void*)&pp);
		}
		al = &m_cellformat[index].alignment;
		switch (prop) {
		case MY_ALIGN_HORIZONTAL:al->horizontal = value; break;
		case MY_ALIGN_VERTICAL:al->vertical = value; break;
		case MY_ALIGN_INDENT:al->indent = value; break;
		case MY_ALIGN_JUSTIFYLASTLINE:al->justifylastline=value; break;
		case MY_ALIGN_READINGORDER:al->readingorder=value; break;
		case MY_ALIGN_RELATIVEINDENT:al->relativeindent=value; break;
		case MY_ALIGN_TEXTROTATION:al->textrotation = value; break;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_BORDERINDEX: {
		XLBORDERSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].borderindex;
		else
			indexf = 0;
		if (indexf>0 && indexf < m_bordercount)
			memcpy((void*)&p, (void*)&m_borders[indexf], sizeof(XLBORDERSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLBORDERSTRUCT));
		switch (prop) {
		case MY_BORDER_LEFT:p.left.style = value; break;
		case MY_BORDER_RIGHT:p.right.style=value; break;
		case MY_BORDER_TOP:p.top.style = value; break;
		case MY_BORDER_BOTTOM:p.bottom.style = value; break;
		case MY_BORDER_VERTICAL:p.vertical.style = value; break;
		case MY_BORDER_HORIZONTAL:p.horizontal.style = value; break;
		case MY_BORDER_DIAGONALUP:p.diagonalup = 1; p.diagonal.style = value; break;
		case MY_BORDER_DIAGONALDOWN:p.diagonaldown = 1; p.diagonal.style = value; break;
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findborder((void*)&p);
		if (indexf < 0)indexf = createborder((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.borderindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].borderindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}

	default:break;
	}
	return 0;
}

int32_t XLDocument1::setdoublestyle(int32_t index, int32_t type, int32_t prop, double value)
{
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID: {
		XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].numberformatid;
		else
			indexf = 0;
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.numberformatid = value;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].numberformatid = value;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		XLFONTSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].fontindex;
		else
			indexf = 0;
		if (indexf > 0 && indexf < m_fontcount)
			memcpy((void*)&p, (void*)&m_fonts[indexf], sizeof(XLFONTSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLFONTSTRUCT));
		switch (prop) {
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findfont((void*)&p);
		if (indexf < 0)indexf = createfont((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.fontindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].fontindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_FILLINDEX: {
		XLFILLSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].fillindex;
		else
			indexf = 0;
		if (indexf > 0 && indexf < m_fillcount)
			memcpy((void*)&p, (void*)&m_fills[indexf], sizeof(XLFILLSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLFILLSTRUCT));
		switch (prop) {
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findfill((void*)&p);
		if (indexf < 0)indexf = createfill((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.fillindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].fillindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_ALIGNMENT: {
		XLCELLFORMATSTRUCT pp; XLALIGNSTRUCT* al;
		if (!index) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			index = createcellformat((void*)&pp);
		}
		al = &m_cellformat[index].alignment;
		switch (prop) {
		default:break;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_BORDERINDEX: {
		XLBORDERSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].borderindex;
		else
			indexf = 0;
		if (indexf > 0 && indexf < m_bordercount)
			memcpy((void*)&p, (void*)&m_borders[indexf], sizeof(XLBORDERSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLBORDERSTRUCT));
		switch (prop) {
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findborder((void*)&p);
		if (indexf < 0)indexf = createborder((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.borderindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].borderindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}

	default:break;
	}
	return 0;
}

int32_t XLDocument1::setcharstyle(int32_t index, int32_t type, int32_t prop, std::string value)
{
	switch (type) {
	case MY_XLCELLFORMAT_NUMBERFORMATID: {
		int indexf;
		indexf = findnumberformatem(value.data());
		if (indexf < 0)indexf = findnumberformat(value.data());
		if (indexf < 0)indexf = createnumberformat(value.data());
		if (indexf < 0)return 0;
		return setintstyle(index, type,MY_NUMBERFORMAT_ID , indexf);
	}
	case MY_XLCELLFORMAT_ALIGNMENT: {
		switch (prop) {
		case MY_ALIGN_HORIZONTAL:
		case MY_ALIGN_VERTICAL:
			return setintstyle(index, type, prop, XLAlignmentStyleFromString(value));
		case MY_ALIGN_INDENT:return 0;
		case MY_ALIGN_JUSTIFYLASTLINE:return 0;
		case MY_ALIGN_READINGORDER:return 0;
		case MY_ALIGN_RELATIVEINDENT:return 0;
		case MY_ALIGN_TEXTROTATION:return 0;
		case MY_ALIGN_SHRINKTOFIT:return 0;
		case MY_ALIGN_WRAPTEXT:return 0;
		default:break;
		}
		return 0;
	}
	case MY_XLCELLFORMAT_FONTINDEX: {
		XLFONTSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].fontindex;
		else
			indexf = 0;
		if (indexf && indexf < m_fontcount)
			memcpy((void*)&p, (void*)&m_fonts[indexf], sizeof(XLFONTSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLFONTSTRUCT));
		switch (prop) {
		case MY_XLFONT_NAME: {
			if (value.length()) {
				auto len = value.length();
				if (len < sizeof(p.name) - 1)strcpy(p.name, value.data());
			}
			break;
		case MY_XLFONT_COLOR: {
			if (value.length()) {
				int rgb = nametorgb(value.data());
				if (rgb != -1) {
					p.fg.color.alpha = 0;
					memcpy(&p.fg.color.red, &rgb, 3);
				}
				else {
					XLColor c(value);
					p.fg.color.alpha = c.alpha();
					p.fg.color.red = c.red();
					p.fg.color.green = c.green();
					p.fg.color.blue = c.blue();
				}
				p.hascolor = 1;
			}
			break;
		}
		case MY_XLFONT_UNDERLINE:
			return setintstyle(index, type, prop, XLUnderlineStyleFromString(value));
		}
		default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findfont((void*)&p);
		if (indexf < 0)indexf = createfont((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.fontindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].fontindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save=1;
		return index;
	}
	case MY_XLCELLFORMAT_BORDERINDEX: {
		XLBORDERSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].borderindex;
		else
			indexf = 0;
		if (indexf && indexf < m_bordercount)
			memcpy((void*)&p, (void*)&m_borders[indexf], sizeof(XLBORDERSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLBORDERSTRUCT));
		switch (prop) {
		case MY_BORDER_LEFT_COLOR:
			if (value.length()) {
				int rgb = nametorgb(value.data());
				if (rgb != -1) {
					p.left.color.argb.alpha = 0;
					memcpy(&p.left.color.argb.red, &rgb, 3);
				}
				else {
					XLColor c(value);
					p.left.color.argb.alpha = c.alpha();
					p.left.color.argb.red = c.red();
					p.left.color.argb.green = c.green();
					p.left.color.argb.blue = c.blue();
				}
				p.left.hascolor = 1;
			}
			break;
		case MY_BORDER_RIGHT_COLOR:
			if (value.length()) {
				int rgb = nametorgb(value.data());
				if (rgb != -1) {
					p.right.color.argb.alpha = 0;
					memcpy(&p.right.color.argb.red, &rgb, 3);
				}
				else {
					XLColor c(value);
					p.right.color.argb.alpha = c.alpha();
					p.right.color.argb.red = c.red();
					p.right.color.argb.green = c.green();
					p.right.color.argb.blue = c.blue();
				}
				p.right.hascolor = 1;
			}
			break;
		case MY_BORDER_TOP_COLOR:
			if (value.length()) {
				int rgb = nametorgb(value.data());
				if (rgb != -1) {
					p.top.color.argb.alpha = 0;
					memcpy(&p.top.color.argb.red, &rgb, 3);
				}
				else {
					XLColor c(value);
					p.top.color.argb.alpha = c.alpha();
					p.top.color.argb.red = c.red();
					p.top.color.argb.green = c.green();
					p.top.color.argb.blue = c.blue();
				}
				p.top.hascolor = 1;
			}
			break;
		case MY_BORDER_BOTTOM_COLOR:
			if (value.length()) {
				int rgb = nametorgb(value.data());
				if (rgb != -1) {
					p.bottom.color.argb.alpha = 0;
					memcpy(&p.bottom.color.argb.red, &rgb, 3);
				}
				else {
					XLColor c(value);
					p.bottom.color.argb.alpha = c.alpha();
					p.bottom.color.argb.red = c.red();
					p.bottom.color.argb.green = c.green();
					p.bottom.color.argb.blue = c.blue();
				}
				p.bottom.hascolor = 1;
			}
			break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findborder((void*)&p);
		if (indexf < 0)indexf = createborder((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.borderindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].borderindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}
	case MY_XLCELLFORMAT_FILLINDEX: {
		XLFILLSTRUCT p; XLCELLFORMATSTRUCT pp; int indexf; int c;
		if (m_cellformatcount)
			indexf = m_cellformat[index].fillindex;
		else
			indexf = 0;
		if (indexf && indexf < m_fillcount)
			memcpy((void*)&p, (void*)&m_fills[indexf], sizeof(XLFILLSTRUCT));
		else
			memset((void*)&p, 0, sizeof(XLFILLSTRUCT));
		switch (prop) {
			case MY_FILL_COLOR: {
				if (value.length()) {
					int rgb = nametorgb(value.data());
					if (rgb != -1) {
						p.fg.color.alpha = 0;
						memcpy(&p.fg.color.red, &rgb, 3);
					}
					else {
						XLColor c(value);
						p.fg.color.alpha = c.alpha();
						p.fg.color.red = c.red();
						p.fg.color.green = c.green();
						p.fg.color.blue = c.blue();
					}
					p.hasfgcolor = 1;
					p.filltype = XLPatternFill;
				}
				break;
			}
			case MY_FILL_BACKGROUNDCOLOR: {
				if (value.length()) {
					int rgb = nametorgb(value.data());
					if (rgb != -1) {
						p.fg.color.alpha = 0;
						memcpy(&p.fg.color.red, &rgb, 3);
					}
					else {
						XLColor c(value);
						p.bg.color.alpha = c.alpha();
						p.bg.color.red = c.red();
						p.bg.color.green = c.green();
						p.bg.color.blue = c.blue();
					}
					p.hasbgcolor = 1;
					p.filltype = XLPatternFill;
				}
				break;
			}
			default:break;
		}
		if (index)
			c = countcellformat(type, indexf);
		else
			c = 2;
		indexf = findfill((void*)&p);
		if (indexf < 0)indexf = createfill((void*)&p);
		if (c > 1) {
			memcpy((void*)&pp, &m_cellformat[index], sizeof(XLCELLFORMATSTRUCT));
			pp.fillindex = indexf;
			index = createcellformat((void*)&pp);
		}
		else {
			m_cellformat[index].fillindex = indexf;
			m_cellformat[index].unsave = 1;
		}
		m_save = 1;
		return index;
	}
	default:break;
	}
	return 0;
}

XLDocument * XLDocument1::doc()
{
	return m_doc;
}

void XLDocument1::create(const std::string& fileName, bool forceOverwrite)
{
	m_doc->create(fileName, forceOverwrite);
	getallstyles();
};
void XLDocument1::open(const std::string& fileName)
{
	m_doc->open(fileName);
	getallstyles();
};
void XLDocument1::save()
{
	setallstyles();
	m_doc->save();
};

void XLDocument1::close()
{
	m_doc->close();
}

XLWorkbook1 XLDocument1::workbook()
{
	XLWorkbook1 m_wb1=XLWorkbook1(this,m_doc->workbook());
	return m_wb1;
}

//---------------class XLWorkbook1------------------------------------------------------------------------

XLWorkbook1::XLWorkbook1(XLDocument1 * doc1,const XLWorkbook &wb)
{
	m_doc1 = doc1;
	m_wb = wb;
}

XLWorksheet1 XLWorkbook1::worksheet(uint16_t n)
{
	const XLWorksheet & ws=m_wb.worksheet(n);
	return XLWorksheet1(m_doc1,ws);
}

XLWorksheet1 XLWorkbook1::worksheet(const std::string & name)
{
	const XLWorksheet& ws = m_wb.worksheet(name);
	return XLWorksheet1(m_doc1,ws);

}

void XLWorkbook1::addWorksheet(const std::string& name)
{
	m_wb.addWorksheet(name);
}
void XLWorkbook1::cloneSheet(const std::string& name, const std::string& newname)
{
	m_wb.cloneSheet(name,newname);
}
void XLWorkbook1::deleteSheet(const std::string& name)
{
	m_wb.deleteSheet(name);
}
unsigned int XLWorkbook1::worksheetCount()
{
	return m_wb.worksheetCount();
}

XLWorksheet1 XLWorkbook1::worksheet(const char *name)
{
	const XLWorksheet& ws = m_wb.worksheet((const std::string&)std::string(name));
	return XLWorksheet1(m_doc1,ws);
}

//---------------------class XLWorksheet1-----------------------------------------------------------------

XLWorksheet1::XLWorksheet1(XLDocument1 *doc1, const XLWorksheet & ws)
{
	m_doc1 = doc1;
	m_ws = ws;
	m_index = ws.index();
#ifdef MY_DRAWING
	XLDrawing1();
#endif
}

XLCell1 XLWorksheet1::cell(const std::string &address)
{
	return XLCell1(m_doc1,this ,m_ws.cell(address) );
}

XLCell1 XLWorksheet1::cell(char * address)
{
	return XLCell1(m_doc1, this, m_ws.cell(_strupr(address)));
}

XLCell1 XLWorksheet1::cell(int32_t row,int16_t column)
{
	return XLCell1(m_doc1,this,m_ws.cell(row,column));
}

XLCellRange1 XLWorksheet1::range()
{
	return XLCellRange1(m_doc1, this, m_ws.range());
}

XLCellRange1 XLWorksheet1::range(std::string const& address)
{
	return XLCellRange1(m_doc1,this,m_ws.range(address));
}

XLCellRange1 XLWorksheet1::range(char *address)
{
	return XLCellRange1(m_doc1, this, m_ws.range(_strupr(address)));
}

void XLWorksheet1::setSelected(bool sel)
{
	m_ws.setSelected(sel);
}
void XLWorksheet1::merge(const std::string &address)
{
	m_ws.merges().appendMerge(address);
}

int16_t XLWorksheet1::columnCount()
{
	return m_ws.columnCount();
}
int32_t XLWorksheet1::rowCount()
{
	return m_ws.rowCount();
}

XLCellReference XLWorksheet1::lastCell()
{
	return m_ws.lastCell();
}

void XLWorksheet1::copyRange(XLRECT* from, XLRECT* to)
{
	int row_from, row_to, col_from, col_to;
	
	m_doc1->setallstyles();

	for (row_from = from->bottom, row_to = to->bottom; row_from >= from->top; row_from--, row_to--) {

//		h = m_ws.row(row_from).height();
//		m_ws.row(row_to).setHeight(h);

		for (col_from = from->left, col_to = to->left; col_from <= from->right; col_from++, col_to++) {
			XLCell c_to = m_ws.cell(row_to, col_to);
			XLCellReference cr_to = c_to.cellReference();
			if (hasMerges()) {
				int32_t mc = m_ws.merges().findMergeByCell(cr_to);
				if (mc >= 0) {
					m_ws.merges().deleteMerge(mc);
				}
			}
		}
		for (col_from = from->left, col_to = to->left; col_from <= from->right; col_from++, col_to++) {
			XLCell c_from = m_ws.cell(row_from, col_from);
			XLCell c_to = m_ws.cell(row_to, col_to);
			XLCellReference cr_from = c_from.cellReference();
			if (hasMerges()) {
				int32_t mc = m_ws.merges().findMergeByCell(cr_from);
				if (mc >= 0) {
					const char* s = m_ws.merges().merge(mc);
					auto ss = std::string(s);
					auto pos = ss.find(":");
					if (pos) {
						XLCellReference cr_beg = XLCellReference(ss.substr(0, pos));
						XLCellReference cr_end = XLCellReference(ss.substr(pos + 1));
						cr_beg.setRow(row_to);
						cr_end.setRow(row_to);
						mc = m_ws.merges().findMergeByCell(cr_beg);
						if (mc >= 0) {
							m_ws.merges().deleteMerge(mc);
						}
						m_ws.merges().appendMerge(cr_beg.address() + ":" + cr_end.address());
					}
				}
			}
			c_to.copyFrom((const XLCell)c_from);
		}
	}
}
int XLWorksheet1::addPicture(void* buffer, int bufferlen, char* ext, XLRECT* rect)
{
	return m_doc1->appendPictures(m_index, buffer, bufferlen, ext, rect);
}

int XLWorksheet1::addPicture(std::string name,XLRECT* rect)
{
	struct stat st; int ret = -1; char* ext;
	char* pic = name.data();
	int pos=name.find_last_of(".");
	if (pos == std::_Meta_npos)ext = (char *)"png";
	else ext = pic + pos + 1;
	stat(pic, &st);
	int buflen = st.st_size;
	char *buf = (char*)malloc(buflen);
	if (buf) {
		FILE* fi = fopen(pic, (const char*)"rb");
		fread((void*)buf, buflen, 1, fi);
		fclose(fi);
		ret=m_doc1->appendPictures(m_index, buf, buflen, ext, rect);
		free(buf);
	}
	return ret;
}

//-------------------class XLCell1---------------------------------------------------------

XLCell1::XLCell1(XLDocument1 *doc1,XLWorksheet1 * ws1,const XLCell & c)
{
	m_doc1 = doc1;
	m_ws1 = ws1;
	m_c = c;
}

XLWorksheet1* XLCell1::ws1() { return m_ws1; };

void XLCell1::copyFrom(XLCell1 c1)
{
	int16_t fromsheetno = c1.ws1()->index();
	int32_t fromrow = c1.c().cellReference().row();
	int16_t fromcol = c1.c().cellReference().column();
	int32_t fromindex = m_doc1->findcharacter(fromsheetno,fromrow,fromcol);
	m_c.copyFrom(c1.c());
	if (fromindex >= 0) {
		int16_t sheetno = m_ws1->index();
		int32_t row = m_c.cellReference().row();
		int16_t col = m_c.cellReference().column();
		m_doc1->copycharacter(fromindex, sheetno, row, col);
	}
}

XLFont1 XLCell1::font()
{
	return XLFont1(m_doc1,this);
}

XLFill1 XLCell1::fill()
{
	return XLFill1(m_doc1,this);
}

XLBorders1 XLCell1::borders()
{
	return XLBorders1(m_doc1,this);
}

XLCharacters1 XLCell1::characters(int16_t start, int16_t len)
{
	return XLCharacters1(m_doc1, this, start, len);
}

XLBorder1 XLCell1::borders(const int32_t index)
{
	return XLBorders1(m_doc1, this).item(index);
}

XLCellValueProxy& XLCell1::value() {
	m_doc1->setcharacters();
	return m_c.value();
}

int32_t XLCell1::horizontalAlignment()
{
	return m_doc1->getintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL);
}

XLCell1 & XLCell1::setHorizontalAlignment(int32_t value)
{
	auto index = m_doc1->setintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL, value);
	m_c.setCellFormat(index);
	return *this;
}

XLCell1 & XLCell1::setHorizontalAlignment(std::string value)
{
	auto index = m_doc1->setcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL, value);
	m_c.setCellFormat(index);
	return *this;
}
int32_t XLCell1::verticalAlignment()
{
	return m_doc1->getintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL);
}

XLCell1 & XLCell1::setVerticalAlignment(int32_t value)
{
	auto index = m_doc1->setintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL, value);
	m_c.setCellFormat(index);
	return *this;
}

XLCell1 & XLCell1::setVerticalAlignment(std::string value)
{
	auto index = m_doc1->setcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL, value);
	m_c.setCellFormat(index);
	return *this;
}

bool XLCell1::wraptext()
{
	return m_doc1->getboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_WRAPTEXT);
}

XLCell1 & XLCell1::setWraptext(bool value)
{
	auto index = m_doc1->setboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_WRAPTEXT, value);
	m_c.setCellFormat(index);
	return *this;
}

bool XLCell1::shrinktofit()
{
	return m_doc1->getboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_SHRINKTOFIT);
}

XLCell1 & XLCell1::setShrinktofit(bool value)
{
	auto index = m_doc1->setboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_SHRINKTOFIT, value);
	m_c.setCellFormat(index);
	return *this;
}
char * XLCell1::numberFormat()
{
	return m_doc1->getcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_NUMBERFORMATID, MY_NUMBERFORMAT_CODE);
}

XLCell1 & XLCell1::setNumberFormat(std::string value)
{
	auto index = m_doc1->setcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_NUMBERFORMATID, MY_NUMBERFORMAT_CODE, value);
	m_c.setCellFormat(index);
	return *this;
}

//--------------------class XLCharacters1--------------------------------------------------------------

XLCharacters1::XLCharacters1(XLDocument1* doc1, XLCell1 *c1 , int16_t start, int16_t len)
{
	m_doc1 = doc1;
	m_c1 = c1;
	m_start = start;
	m_len = len;
}

XLFont1 XLCharacters1::font()
{
	return XLFont1(m_doc1,this);
}

//------------------class XLCellRange1--------------------------------------------------

XLCellRange1::XLCellRange1(XLDocument1* doc1, XLWorksheet1 * ws1, const XLCellRange & cr)
{
	m_doc1 = doc1;
	m_cr = cr;
	m_ws1 = ws1;
}

void XLCellRange1::rect(XLRECT *rect)
{	
	const XLCellReference tl=m_cr.topLeft();
	const XLCellReference br = m_cr.bottomRight();
	rect->left = tl.column();
	rect->top = tl.row();
	rect->right = br.column();
	rect->bottom = br.row();
}
/*
const XLCellValueProxy& XLCellRange1::value()
{
	XLRECT rc;
	rect(&rc);
	XLWorksheet1 * s1 = ws1();
	const XLCell1 & c = s1->cell((int32_t)rc.top, (int16_t)rc.left);
	return c.value();
}
*/
std::string XLCellRange1::address()
{
	return m_cr.address();
}

XLBordersR1 XLCellRange1::borders()
{
	return XLBordersR1(m_doc1, this);
}


XLBorderR1 XLCellRange1::borders(int32_t index)
{
	return XLBordersR1(m_doc1,this).item(index);
}

XLFont1 XLCellRange1::font()
{
	return XLFont1(m_doc1,this);
}

XLWorksheet1* XLCellRange1::ws1() { return m_ws1; };

XLFill1 XLCellRange1::fill()
{
	return XLFill1(m_doc1,this);
}

void XLCellRange1::merge()
{
	m_ws1->merge(m_cr.address());
}

void XLCellRange1::copyFrom(std::string address)
{
	XLRECT from, to;
	XLWorksheet ws = m_ws1->ws();

	to.right=m_cr.bottomRight().column();
	to.bottom = m_cr.bottomRight().row();
	to.left = m_cr.topLeft().column();
	to.top = m_cr.topLeft().row();

	XLCell c=ws.cell(address);
	from.left = c.cellReference().column();
	from.top= c.cellReference().row();
	from.right = from.left + to.right - to.left;
	from.bottom = from.top + to.bottom - to.top;
	m_ws1->copyRange(&from, &to);
}

void XLCellRange1::copyTo(std::string address)
{
	XLRECT from, to;
	XLWorksheet ws = m_ws1->ws();

	from.right = m_cr.bottomRight().column();
	from.bottom = m_cr.bottomRight().row();
	from.left = m_cr.topLeft().column();
	from.top = m_cr.topLeft().row();

	XLCell c = ws.cell(address);
	to.left = c.cellReference().column();
	to.top = c.cellReference().row();
	to.right = to.left + from.right - from.left;
	to.bottom = to.top + from.bottom - from.top;
	m_ws1->copyRange(&from, &to);
}

void XLCellRange1::insert()
{
	XLRECT from, to;
	from.top = m_cr.topLeft().row();
	from.left = 1;
	to.top = from.top + 1;
	to.left = 1;
	XLWorksheet ws = m_ws1->ws();
	from.bottom = ws.rowCount();
	to.bottom = from.bottom + 1;
	from.right = from.left + ws.columnCount() - 1;
	to.right = from.right;

	m_ws1->copyRange(&from,&to);
}

void XLCellRange1::setpropchar(int32_t type, int32_t prop, std::string value)
{
	int32_t index;
	for (auto it = m_cr.begin(); it != m_cr.end(); ++it) {
		index = m_doc1->setcharstyle(it->cellFormat(), type, prop, value);
		it->setCellFormat(index);
	}
}

void XLCellRange1::setpropdouble(int32_t type, int32_t prop,double value)
{
	int32_t index;
	for (auto it = m_cr.begin(); it != m_cr.end(); ++it) {
		index = m_doc1->setdoublestyle(it->cellFormat(), type, prop, value);
		it->setCellFormat(index);
	}
}


void XLCellRange1::setpropint(int32_t type, int32_t prop, int32_t value)
{
	int32_t index;
	for (auto it = m_cr.begin(); it != m_cr.end(); ++it) {
		index = m_doc1->setintstyle(it->cellFormat(), type, prop, value);
		it->setCellFormat(index);
	}
}

void XLCellRange1::setpropbool(int32_t type, int32_t prop, bool value)
{
	int32_t index;
	for (auto it = m_cr.begin(); it != m_cr.end(); ++it) {
		index = m_doc1->setboolstyle(it->cellFormat(), type, prop, value);
		it->setCellFormat(index);
	}
}

void XLCellRange1::setHorizontalAlignment(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL, value);
}

void XLCellRange1::setHorizontalAlignment(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL, value);
}

void XLCellRange1::setVerticalAlignment(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL, value);
}

void XLCellRange1::setVerticalAlignment(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL, value);
}
void XLCellRange1::setWraptext(bool value)
{
	setpropbool(MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_WRAPTEXT, value);
}

void XLCellRange1::setShrinktofit(bool value)
{
	setpropbool(MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_SHRINKTOFIT, value);
}

void XLCellRange1::setNumberFormat(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_NUMBERFORMATID, MY_NUMBERFORMAT_CODE, value);
}

//--------------------class XLBorders1--------------------------------------------------------

XLBorders1::XLBorders1(XLDocument1* doc1, XLCell1 * c1)
{
	m_doc1 = doc1;
	m_c1 = c1;
}


XLBordersR1::XLBordersR1(XLDocument1 *doc1,XLCellRange1 * cr1)
{
	m_doc1 = doc1;
	m_cr1 = cr1;
}

XLBorder1 XLBorders1::item(int32_t index)
{
	return XLBorder1(m_doc1,*this,index);
}

XLBorderR1 XLBordersR1::item(int32_t index)
{
	return XLBorderR1(m_doc1, *this, index);
}


//-----------------class XLBorder1--------------------------------------------------

XLBorder1::XLBorder1(XLDocument1 *doc1,const XLBorders1 & bs1,int32_t index)
{
	m_doc1 = doc1;
	m_bs1 = bs1;
	m_index = index;
}

XLBorderR1::XLBorderR1(XLDocument1* doc1, const XLBordersR1 & bs1, int32_t index)
{
	m_doc1 = doc1;
	m_bs1 = bs1;
	m_index = index;
}

int32_t XLBorder1::lineStyle()
{
	XLStyleIndex cf;
	XLCell c = m_bs1.c1()->c();
	cf = c.cellFormat();
	return m_doc1->getintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index);
}

int32_t XLBorderR1::lineStyle()
{
	XLRECT rect;
	XLCellRange1* cr1 = m_bs1.cr1();
	cr1->rect(&rect);
	XLWorksheet1* s1 = cr1->ws1();
	XLCell1 c1 = s1->cell((int32_t)rect.top, (int16_t)rect.left);
	XLBorders1 bs1 = c1.borders();
	XLBorder1  b1 = bs1.item(0);
	return b1.lineStyle();
}

void XLBorder1::setLineStyle(int32_t ls)
{
	XLStyleIndex cf;
	XLCell1 *c1 = m_bs1.c1();
	XLCell c=c1->c();
	cf = c.cellFormat();
	cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index, ls);
	c.setCellFormat(cf);
}

void  XLBorderR1::setLineStyle(int32_t ls)
{	
	XLStyleIndex cf;
	XLRECT rect;
	XLCellRange1 *cr1 = m_bs1.cr1();
	XLWorksheet1 *s1=cr1->ws1();
	cr1->rect(&rect);
	switch (m_index) {
	case 0:
		for (auto i = rect.top; i <= rect.bottom; i++) {
			XLCell1 c1=s1->cell(i,rect.left);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index, ls);
			c.setCellFormat(cf);
		}
		break;
	case 1:
		for (auto i = rect.top; i <= rect.bottom; i++) {
			XLCell1 c1 = s1->cell(i,rect.right);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index, ls);
			c.setCellFormat(cf);
		}
		break;
	case 2:
		for (auto i = rect.left; i <= rect.right; i++) {
			XLCell1 c1 = s1->cell(rect.top,i);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index, ls);
			c.setCellFormat(cf);
		}
		break;
	case 3:
		for (auto i = rect.left; i <= rect.right; i++) {
			XLCell1 c1 = s1->cell(rect.bottom, i);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index, ls);
			c.setCellFormat(cf);
		}
		break;
	}
}

int32_t XLBorder1::color()
{
	XLStyleIndex cf;
	XLCell1 *c1 = m_bs1.c1();
	XLCell c=c1->c();
	cf = c.cellFormat();
	return m_doc1->getintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index);
}

int32_t XLBorderR1::color()
{
	XLRECT rect;
	XLCellRange1* cr1 = m_bs1.cr1();
	cr1->rect(&rect);
	XLWorksheet1* s1 = cr1->ws1();
	XLCell1 c1 = s1->cell((int32_t)rect.top, (int16_t)rect.left);
	XLBorders1 bs1 = c1.borders();
	XLBorder1  b1 = bs1.item(0);
	return b1.color();
}

void XLBorder1::setColor(std::string color)
{
	XLStyleIndex cf;
	XLCell c = m_bs1.c1()->c();
	cf = c.cellFormat();
	cf = m_doc1->setcharstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index+8,color);
	c.setCellFormat(cf);
}

void XLBorderR1::setColor(std::string color)
{
	XLStyleIndex cf;
	XLRECT rect;
	XLCellRange1* cr1 = m_bs1.cr1();
	XLWorksheet1* s1 = cr1->ws1();
	int colindex = m_index + 8;
	cr1->rect(&rect);
	switch (m_index) {
	case 0:
		for (auto i = rect.top; i <= rect.bottom; i++) {
			XLCell1 c1 = s1->cell(i, rect.left);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setcharstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, colindex, color);
			c.setCellFormat(cf);
		}
		break;
	case 1:
		for (auto i = rect.top; i <= rect.bottom; i++) {
			XLCell1 c1 = s1->cell(i, rect.right);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setcharstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, colindex, color);
			c.setCellFormat(cf);
		}
		break;
	case 2:
		for (auto i = rect.left; i <= rect.right; i++) {
			XLCell1 c1 = s1->cell(rect.top, i);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setcharstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, colindex, color);
			c.setCellFormat(cf);
		}
		break;
	case 3:
		for (auto i = rect.left; i <= rect.right; i++) {
			XLCell1 c1 = s1->cell(rect.bottom, i);
			XLCell c = c1.c();
			cf = c.cellFormat();
			cf = m_doc1->setcharstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, colindex, color);
			c.setCellFormat(cf);
		}
		break;
	}
}

//-------------------class XLFill1----------------------------------------------------
XLFill1::XLFill1(XLDocument1* doc1, XLCell1 * c1)
{
	m_t = 0;
	m_doc1 = doc1;
	m_c1 = c1;
}

XLFill1::XLFill1(XLDocument1* doc1, XLCellRange1 * cr1)
{
	m_t = 1;
	m_doc1 = doc1;
	m_cr1 = cr1;
}

void XLFill1::setpropint(int32_t type, int32_t prop, int32_t value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = c.cellFormat();
		index = m_doc1->setintstyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setintstyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}

}

void XLFill1::setpropdouble(int32_t type, int32_t prop, double value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = c.cellFormat();
		index = m_doc1->setdoublestyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setdoublestyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}

}

void XLFill1::setpropbool(int32_t type, int32_t prop, bool value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = c.cellFormat();
		index = m_doc1->setboolstyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setboolstyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}
}

void XLFill1::setpropchar(int32_t type, int32_t prop, std::string value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = m_doc1->setcharstyle(c.cellFormat(), type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = m_doc1->setcharstyle(it->cellFormat(), type, prop, value);
			it->setCellFormat(index);
		}
	}
}



int32_t XLFill1::color()
{
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FILLINDEX, MY_FILL_COLOR);
}
void XLFill1::setColor(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FILLINDEX, MY_FILL_COLOR, value);
}
void XLFill1::setColor(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_FILLINDEX, MY_FILL_COLOR, value);
}

int32_t XLFill1::backgroundColor()
{
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FILLINDEX, MY_FILL_BACKGROUNDCOLOR);
}

void XLFill1::setBackgroundColor(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FILLINDEX, MY_FILL_BACKGROUNDCOLOR, value);
}
void XLFill1::setBackgroundColor(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_FILLINDEX, MY_FILL_BACKGROUNDCOLOR, value);
}

int32_t XLFill1::patternType()
{
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FILLINDEX, MY_FILL_PATTERNTYPE);
}

void XLFill1::setPatternType(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_FILLINDEX, MY_FILL_PATTERNTYPE, value);
}

//------------------class XLFont1-----------------------------------------------------

XLFont1::XLFont1(XLDocument1 *doc1,XLCell1 * c1)
{
	m_t = 0;
	m_doc1 = doc1;
	m_c1 = c1;
}

XLFont1::XLFont1(XLDocument1* doc1, XLCellRange1 * cr1)
{
	m_t = 1;
	m_doc1 = doc1;
	m_cr1 = cr1;
}

XLFont1::XLFont1(XLDocument1* doc1, XLCharacters1 * ch1)
{
	m_t = 2;
	m_doc1 = doc1;
	m_ch1 = ch1;
}

//XLFont1::~XLFont1(){}

void XLFont1::setpropchar(int32_t type,int32_t prop, std::string value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = m_doc1->setcharstyle(c.cellFormat(), type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = m_doc1->setcharstyle(it->cellFormat(), type, prop, value);
			it->setCellFormat(index);
		}
	}
	if (m_t == 2) {
		int32_t index, indexf;
		XLCHARACTERSTRUCT pp;
		XLCell1 *c1 = m_ch1->c1();
		XLWorksheet1 *ws1 = c1->ws1();
		XLCell c = c1->c();
		pp.sheetno = ws1->index();
		pp.row = c.cellReference().row();
		pp.col = c.cellReference().column();
		pp.start = m_ch1->start();
		pp.len = m_ch1->len();
		pp.indexf = 0;
		index = m_doc1->findcharacter((void*)&pp);
		if (index < 0)index = m_doc1->createcharacter((void*)&pp);
		if (index < 0)return;
		indexf = m_doc1->m_characters[index].indexf;
		indexf = m_doc1->setcharstyle(indexf, type, prop, value);
		m_doc1->m_characters[index].indexf = indexf;
		return;
	}
}

void XLFont1::setpropint(int32_t type, int32_t prop, int32_t value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = c.cellFormat();
		index = m_doc1->setintstyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setintstyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}
	if (m_t == 2) {
		int32_t index, indexf;
		XLCHARACTERSTRUCT pp;
		XLCell1 *c1 = m_ch1->c1();
		XLWorksheet1 *ws1 = c1->ws1();
		XLCell c = c1->c();
		pp.sheetno = ws1->index();
		pp.row = c.cellReference().row();
		pp.col = c.cellReference().column();
		pp.start = m_ch1->start();
		pp.len = m_ch1->len();
		pp.indexf = 0;
		index = m_doc1->findcharacter((void*)&pp);
		if (index < 0)index = m_doc1->createcharacter((void*)&pp);
		if (index < 0)return;
		indexf = m_doc1->m_characters[index].indexf;
		indexf = m_doc1->setintstyle(indexf, type, prop, value);
		m_doc1->m_characters[index].indexf = indexf;
		return;
	}

}

void XLFont1::setpropbool(int32_t type, int32_t prop, bool value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1->c();
		index = c.cellFormat();
		index = m_doc1->setboolstyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1->cr().begin(); it != m_cr1->cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setboolstyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}
	if (m_t == 2) {
		int32_t index, indexf;
		XLCHARACTERSTRUCT pp;
		XLCell1 *c1 = m_ch1->c1();
		XLWorksheet1 *ws1 = c1->ws1();
		XLCell c = c1->c();
		pp.sheetno =ws1->index();
		pp.row = c.cellReference().row();
		pp.col = c.cellReference().column();
		pp.start = m_ch1->start();
		pp.len = m_ch1->len();
		pp.indexf = 0;
		index = m_doc1->findcharacter((void*)&pp);
		if (index < 0)index = m_doc1->createcharacter((void*)&pp);
		if (index < 0)return;
		indexf = m_doc1->m_characters[index].indexf;
		indexf = m_doc1->setboolstyle(indexf, type, prop, value);
		m_doc1->m_characters[index].indexf = indexf;
		return;
	}
}

char* XLFont1::name() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getcharstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_NAME);
}

XLFont1 XLFont1::setName(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FONTINDEX,MY_XLFONT_NAME,value);
	return *this;
}


bool XLFont1::bold() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getboolstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_BOLD);
}


XLFont1 XLFont1::setBold(bool value)
{
	setpropbool(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_BOLD, value);
	return *this;
}

bool XLFont1::italic() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getboolstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_ITALIC);
}

XLFont1 XLFont1::setItalic(bool value)
{
	setpropbool(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_ITALIC, value);
	return *this;
}

bool XLFont1::strikethrough() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getboolstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_STRIKETHROUGH);
}

XLFont1 XLFont1::setStrikethrough(bool value)
{
	setpropbool(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_STRIKETHROUGH, value);
	return *this;
}

int32_t XLFont1::underline() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_UNDERLINE);
}

XLFont1 XLFont1::setUnderline(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_UNDERLINE, value);
	return *this;
}

XLFont1 XLFont1::setUnderline(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_UNDERLINE, value);
	return *this;
}

int32_t XLFont1::size() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_SIZE);
}

XLFont1 XLFont1::setSize(int32_t value)
{
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_SIZE, value);
	return *this;
}

bool XLFont1::superscript() {
	int n;
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	n=m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN);
	if (n == 2)return true;
	return false;
}

XLFont1 XLFont1::setSuperscript(bool value)
{
	int32_t cfindex = 0; int n;
	if (value)n = 2;
	else n = 0;
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN,n);
	return *this;
}
bool XLFont1::subscript() {
	int n;
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	n = m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN);
	if (n == 1)return true;
	return false;
}

XLFont1 XLFont1::setSubscript(bool value)
{
	int n;
	if (value)n = 1;
	else n = 0;
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN, n);
	return *this;
}

int32_t XLFont1::color() {
	XLCell c = m_c1->c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_COLOR);
}

XLFont1 XLFont1::setColor(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_COLOR, value);
	return *this;
}

int utf8next(utf8_int8_t* str) {
	utf8_int8_t ch = *str; int n;
	if (0xf0 == (0xf8 & ch)) {
		n = 4;
	}
	else if (0xe0 == (0xf0 & ch)) {
		n = 3;
	}
	else if (0xc0 == (0xe0 & ch)) {
		n = 2;
	}
	else {
		n = 1;
	}
	return n;
}

utf8_int8_t* utf8substr(utf8_int8_t* str, int start, int len, int* outlen);

utf8_int8_t* utf8substr(utf8_int8_t* str, int start, int len, int* outlen) {
	utf8_int8_t* t = str; utf8_int8_t* st = NULL; int kk;
	size_t length = 0; size_t n = SIZE_MAX; size_t k;
	k = 0; //if (!len)len = SIZE_MAX;
	while ((size_t)(str - t) < n && '\0' != *str) {
		if (length == start)st = str;
		kk = utf8next(str);
		str += kk;
		if (st)k += kk;
		length++;
		if (st) {
			len--;
			if (len<1)break;
		}
	}
	if (st && !len) {
		*outlen = k;
	}
	else {
		*outlen = 0;
	}
	return (utf8_int8_t*)st;
}

std::string xlrtf(std::string s, int32_t start, int32_t len, std::string rtf)
{
int32_t slen; char* ss; int32_t all = 0;
slen = utf8len((utf8_int8_t*)s.data());
if (!slen)return "";
if (!len)len = slen - start + 1;
if (start >= slen)return "";
if (start + len > slen+1)len = slen - start;
if (len < 1)return "";
ss = (char*)utf8substr((utf8_int8_t*)s.data(), 0, start - 1, &all);
std::string s1 = std::string(ss,all);
ss = (char*)utf8substr((utf8_int8_t*)s.data(), start-1,len, &all);
std::string s2 = std::string(ss, all);
ss = (char*)utf8substr((utf8_int8_t*)s.data(), start - 1+len, slen-len-start+1, &all);
std::string s3 = std::string(ss, all);
std::string out = std::string("");
if (s1.length()) {
	out = out + "<r><t>" + s1 + "</t></r>";
}
if (s2.length()) {
	out = out + "<r><rPr>" + rtf + "</rPr><t>" + s2 + "</t></r>";
}
if (s3.length()) {
	out=out+ "<r><t>"+s3+"</t></r>";
}
return out;
}

