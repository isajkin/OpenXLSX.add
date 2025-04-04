#include "myopenxlsx.h"
#include <pugixml.hpp>
#include <XLUtilities.hpp>
#include "utf8.h"

const std::string ShapeNodeNameDr = "xdr:twoCellAnchor";

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

//----------------------class XLDocument1-----------------------------------
#ifdef MY_XLDRAWING
char* XLDocument1::shapeAttribute(int sheetXmlNo, int shapeNo, char* path)
{
	char* s, * s0; int i, att = 0; XMLNode f; std::string ss, sa;

	XLWorksheet1 wks = workbook().worksheet(sheetXmlNo);
	wks.drawing();

	XLDrawing1& dr = sheetDrawing((uint16_t)sheetXmlNo);
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

void XLDocument1::setShapeAttribute(int sheetXmlNo, int shapeNo, char* path, char* attribute, char* value)
{
	char* s, * s0; int i, att = 0, val = 0; XMLNode f;

	std::string ss, sa, sv;

	XLWorksheet1 wks = workbook().worksheet(sheetXmlNo);
	wks.drawing();

	XLDrawing1& dr = sheetDrawing((uint16_t)sheetXmlNo);
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

#ifdef MY_XMLDATA
XLXmlData* XLDocument1::getXmlData(const std::string& path, bool doNotThrow)
{
	// avoid duplication of code: use const_cast to invoke the const function overload and return a non-const value
	return const_cast<XLXmlData*>(const_cast<XLDocument1 const*>(this)->getXmlData(path, doNotThrow));
}

const XLXmlData* XLDocument1::getXmlData(const std::string& path, bool doNotThrow) const
{
	std::list<XLXmlData>::iterator result = std::find_if(m_doc->m_data.begin(), m_doc->m_data.end(), [&](const XLXmlData& item) { return item.getXmlPath() == path; });
	if (result == m_doc->m_data.end()) {
		if (doNotThrow) return nullptr; // use with caution
		else throw XLInternalError("Path " + path + " does not exist in zip archive.");
	}
	return &*result;
}
#endif

#ifdef MY_XLDRAWING
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
		for (auto& item : m_doc->m_contentTypes.getContentDefItems()) {
			std::string picturesFilename = "xl/media/image" + std::to_string(id) + "." + item.ext();
			if (m_doc->m_archive.hasEntry(picturesFilename)) {
				yes = true;
				break;
			}
		}
		if (!yes) {
			std::string picturesFilename = "xl/media/image" + std::to_string(id) + "." + std::string(ext);
			if (!m_doc->m_archive.hasEntry(picturesFilename)) {
				m_doc->m_archive.addEntry(picturesFilename, v);
				if (!m_doc->m_contentTypes.ExtensionExists(ext))m_doc->m_contentTypes.addDefault(ext, XLContentType::Image);
				break;
			}
		}
		id++;
	}

	std::string drawingsRelsFilename = std::string("xl/drawings/_rels/drawing") + std::to_string(sheetXmlNo) + std::string(".xml.rels");
	m_doc->m_data.emplace_back(this, drawingsRelsFilename);
//	m_doc->m_drwRelationships = XLRelationships(getXmlData(drawingsRelsFilename,false), drawingsRelsFilename);
	constexpr const bool DO_NOT_THROW = true;

	std::string imgtarget = std::string("../media/image") + std::to_string(id) + "." + std::string(ext);
	XLRelationshipItem imgitem = m_doc->m_drwRelationships.addRelationship(XLRelationshipType::Image, imgtarget);

	XLWorksheet1 wks = workbook().worksheet(sheetXmlNo);
	wks.drawing();

	XLDrawing1& dr = sheetDrawing(sheetXmlNo);
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
/*
copy range with height,merge,attribute,value
*/

void XLDocument1::copyRange(int sheetXmlNo, XLRECT* from, XLRECT* to)
{
	int row_from, row_to, col_from, col_to;

	XLWorksheet ws = doc()->workbook().worksheet(sheetXmlNo);

	for (row_from = from->bottom, row_to = to->bottom; row_from >= from->top; row_from--, row_to--) {
		ws.row(row_to).setHeight(ws.row(row_from).height());
		for (col_from = from->left, col_to = to->left; col_from <= from->right; col_from++, col_to++) {
			XLCell c_to = ws.cell(row_to, col_to);
			XLCellReference cr_to = c_to.cellReference();
			int32_t mc = ws.merges().findMergeByCell(cr_to);
			if (mc >= 0) {
				ws.merges().deleteMerge(mc);
			}
		}
		for (col_from = from->left, col_to = to->left; col_from <= from->right; col_from++, col_to++) {
			XLCell c_from = ws.cell(row_from, col_from);
			XLCell c_to = ws.cell(row_to, col_to);
			XLCellReference cr_from = c_from.cellReference();
			int32_t mc = ws.merges().findMergeByCell(cr_from);
			if (mc >= 0) {
				const char* s = ws.merges().merge(mc);
				auto ss = std::string(s);
				auto pos = ss.find(":");
				if (pos) {
					XLCellReference cr_beg = XLCellReference(ss.substr(0, pos));
					XLCellReference cr_end = XLCellReference(ss.substr(pos + 1));
					cr_beg.setRow(row_to);
					cr_end.setRow(row_to);
					mc = ws.merges().findMergeByCell(cr_beg);
					if (mc >= 0) {
						ws.merges().deleteMerge(mc);
					}
					ws.merges().appendMerge(cr_beg.address() + ":" + cr_end.address());
				}
			}
			c_to.copyFrom((const XLCell)c_from);
		}
	}
}


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
XLDocument1::~XLDocument1()
{
	delete m_doc;
};

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
				border->left.style = b.left().style();
				border->right.style = b.bottom().style();
				border->top.style = b.top().style();
				border->horizontal.style = b.horizontal().style();
				border->vertical.style = b.vertical().style();
				border->diagonal.style = b.diagonal().style();
				border->diagonaldown = b.diagonalDown();
				border->diagonalup = b.diagonalUp();

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
				fs->i.color.alpha = f.fontColor().alpha();
				fs->i.color.red = f.fontColor().red();
				fs->i.color.green = f.fontColor().green();
				fs->i.color.blue = f.fontColor().blue();
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
			}
		}
	}
}

void XLDocument1::setallstyles()
{
	int i;
	if (m_save) {
		m_save=0;

		XLNumberFormats& nf = m_doc->styles().numberFormats();
		while (nf.count() < (size_t)m_numberformatcount)nf.create();
		for (i = 0; i < m_numberformatcount; i++) {
			if (m_numberformat[i].unsave) {
				nf[i].setNumberFormatId(m_numberformat[i].id);
				nf[i].setFormatCode(m_numberformat[i].formatcode);
				m_numberformat[i].unsave = 0;
			}
		}
		XLBorders bs = m_doc->styles().borders();
		while (bs.count() < (size_t)m_bordercount) {
			bs.create();
		}
		for (i = 1; i < m_bordercount; i++) {
			XLBORDERSTRUCT* border = m_borders + i;
			if (border->unsave) {
				XLBorder b = bs.borderByIndex(i);
				if(border->bottom.style)b.setBottom((XLLineStyle)border->bottom.style, (XLColor)"00000000", 0);
				if(border->left.style)b.setLeft((XLLineStyle)border->left.style, (XLColor)"00000000", 0);
				if(border->right.style)b.setRight((XLLineStyle)border->right.style, (XLColor)"00000000", 0);
				if(border->top.style)b.setTop((XLLineStyle)border->top.style, (XLColor)"00000000", 0);
				if (border->horizontal.style)b.setHorizontal((XLLineStyle)border->horizontal.style, (XLColor)"00000000", 0);
				if (border->vertical.style)b.setVertical((XLLineStyle)border->vertical.style, (XLColor)"00000000", 0);
				if (border->diagonal.style)b.setDiagonal((XLLineStyle)border->diagonal.style, (XLColor)"00000000", 0);
				b.setDiagonalDown(border->diagonaldown);
				b.setDiagonalUp(border->diagonalup);
				border->unsave = 0;
			}
		}
		XLFonts& fnts = m_doc->styles().fonts();
		while (fnts.count() < (size_t)m_fontcount)fnts.create();
		for (i = 1; i < m_fontcount; i++) {
			XLFONTSTRUCT *f = m_fonts+i; XLFont fs = fnts[i];
			if (f->unsave) {
				if (f->bold)fs.setBold(f->bold);
				if (f->italic)fs.setItalic(f->italic);
				if (f->name[0])fs.setFontName(f->name);
				if (f->size)fs.setFontSize(f->size);
				if (f->charset)fs.setFontCharset(f->charset);
				if (f->family)fs.setFontFamily(f->family);

				if (f->i.rgb) {
					XLColor c(f->i.color.alpha, f->i.color.red, f->i.color.green, f->i.color.blue);
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
		}
		XLCellFormats cf = m_doc->styles().cellFormats();
		while (cf.count() < (size_t)m_cellformatcount)cf.create();
		for (i = 1; i < m_cellformatcount; i++) {
			XLCellFormat c = cf[i];
			XLCELLFORMATSTRUCT* ce = m_cellformat + i;
			if (ce->unsave) {
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
		}
		for (i = 0; i < m_charactercount; i++) {
			XLCHARACTERSTRUCT* cs = m_characters+i;
			int indexf = cs->indexf;
			XLCELLFORMATSTRUCT* cf = m_cellformat + indexf;
			indexf=cf->fontindex;
			XLFONTSTRUCT *f = m_fonts+indexf;

			std::string rtf = "";
			if (f->bold)rtf = rtf + "<b/>";
			if (f->italic)rtf = rtf + "<i/>";
			if (f->underline) {
				if (f->underline == 1) {
					rtf = rtf + "<u/>";
				}
				else {
					if (f->underline == 2) {
						rtf=rtf+"<u val=\"double\"/>";
					}
				}
			}
			if (f->strikethrough)rtf = rtf + "<strike/>";
			if (f->vertalign == 1)rtf = rtf + "<vertAlign val=\"subscript\"/>";
			if (f->vertalign == 2)rtf = rtf + "<vertAlign val=\"superscript\"/>";
			if (f->i.rgb) {
				XLColor color(f->i.color.alpha, f->i.color.red, f->i.color.green, f->i.color.blue);
				color.hex();
				rtf = rtf + "<color rgb=\"" + color.hex() + "\"/>";
			}
			if (f->size) {
				char buf[32];
				char *s=_itoa(f->size, buf, 10);
				rtf = rtf + "<sz val=\""+std::string(s)+"\"/>";
			}
			if (f->charset) {
				char buf[32];
				itoa(f->charset, buf, 10);
				rtf = rtf + "<charset val=\"" + std::string(buf) + "\"/>";
			}
			if(f->name[0])rtf = rtf+"<rFont val=\""+std::string(f->name)+"\"/>";

			if (rtf.length()) {
				XLCell cell = m_doc->workbook().worksheet(cs->sheetno).cell(cs->row, cs->col);
				std::string v = cell.value().getString();
				v = xlrtf(v, cs->start, cs->len, rtf);
				cell.value() = v;
			}
		}
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
		case MY_XLFONT_COLOR:return f->i.rgb;
		default:return 0;
		}
	}
	case MY_XLCELLFORMAT_FILLINDEX:break;
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
		}
		return 0;
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

int32_t XLDocument1::findcharacter(void* p)
{
	int i;
	for (i = 0; i < m_charactercount; i++) {
		if (!memcmp((void*)&m_characters[i], (void*)p, sizeof(XLCHARACTERSTRUCT)-sizeof(int32_t)))return i;
	}
	return -1;
}

int32_t XLDocument1::createcharacter(void* p) {
	m_characters = (XLCHARACTERSTRUCT*)realloc((void*)m_characters, sizeof(XLCHARACTERSTRUCT) * (m_charactercount + 1));
	if (m_characters) {
		memcpy((void*)&m_characters[m_charactercount], p, sizeof(XLCHARACTERSTRUCT));
		m_charactercount++;
		return m_charactercount - 1;
	}
	return -1;
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
		case MY_XLFONT_COLOR:p.i.rgb = value; break;
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
				XLColor c(value);
				p.i.color.alpha = c.alpha();
				p.i.color.red = c.red();
				p.i.color.green = c.green();
				p.i.color.blue = c.blue();
			}
			break;
		}
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
	default:break;
	}
	return 0;
}

XLDocument *XLDocument1::doc()
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
	XLWorkbook1 *wb=new XLWorkbook1(this, (const XLWorkbook) m_doc->workbook());
	return *wb;
}

#ifdef MY_DRAWING
XLDrawing1& XLDocument1::sheetDrawing(uint16_t sheetXmlNo)
{
	using namespace std::literals::string_literals;
	std::string drawingFilename = "xl/drawings/drawing"s + std::to_string(sheetXmlNo) + ".xml"s;

	if (!m_doc->m_archive.hasEntry(drawingFilename)) {
		// ===== Create the sheet drawing file within the archive
		m_doc->m_archive.addEntry(drawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");  // empty XML file, class constructor will do the rest
		if (!m_doc->m_contentTypes.PartNameExists("/" + drawingFilename))
			m_doc->m_contentTypes.addOverride("/" + drawingFilename, XLContentType::Drawing);                          // add content types entry
	}
	constexpr const bool DO_NOT_THROW = true;
	XLXmlData* xmlData = getXmlData(drawingFilename, DO_NOT_THROW);
	if (xmlData == nullptr) // if not yet managed: add the sheet drawing file to the managed files
		xmlData = &m_doc->m_data.emplace_back(this, drawingFilename, "", XLContentType::Drawing);

	return (XLDrawing1&)XLDrawing1(xmlData);
}

bool XLDocument1::hasSheetDrawing(uint16_t sheetXmlNo) const
{
	using namespace std::literals::string_literals;
	return m_doc->m_archive.hasEntry("xl/drawings/drawing"s + std::to_string(sheetXmlNo) + ".xml"s);
}
#endif
//---------------class XLWorkbook1------------------------------------------------------------------------

XLWorkbook1::XLWorkbook1(XLDocument1 *doc1,const XLWorkbook wb)
{
	m_doc1 = doc1;
	m_wb = (XLWorkbook)wb;
}

XLWorkbook1::~XLWorkbook1()
{
}

XLWorksheet1 XLWorkbook1::worksheet(int16_t n)
{
	return XLWorksheet1(m_doc1,m_wb.worksheet(n));
}

XLWorksheet1 XLWorkbook1::worksheet(std::string name)
{
	return XLWorksheet1(m_doc1, m_wb.worksheet(name));

}

//---------------------class XLWorksheet1-----------------------------------------------------------------

XLWorksheet1::~XLWorksheet1() {};

XLWorksheet1::XLWorksheet1() {};

XLWorksheet1::XLWorksheet1(XLDocument1* doc1, XLWorksheet ws)
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
	return XLCell1(m_doc1,*this,m_ws.cell(address));
}

XLCell1 XLWorksheet1::cell(int32_t row,int16_t column)
{
	return XLCell1(m_doc1,*this,m_ws.cell(row,column));
}

XLCellRange1 XLWorksheet1::range(const std::string& address)
{
	return XLCellRange1(m_doc1,*this,m_ws.range(address));
}

XLColumn XLWorksheet1::column(int16_t column) {
	return m_ws.column(column);
}

XLRow XLWorksheet1::row(int32_t row) {
	return m_ws.row(row);
}

void XLWorksheet1::setSelected(bool sel)
{
	m_ws.setSelected(sel);
}
void XLWorksheet1::merge(const std::string &address)
{
	m_ws.merges().appendMerge(address);
}

#ifdef MY_DRAWING
bool XLWorksheet1::hasDrawing() const
{
	return m_doc1->hasSheetDrawing(index());
}

XLDrawing1& XLWorksheet1::drawing()
{
	if (!m_drawing.valid()) {
		// ===== Append xdr namespace attribute to worksheet if not present

		XMLNode docElement = xmlDocument().document_element();
		XMLAttribute xdrNamespace = appendAndGetAttribute(docElement, "xmlns:xdr", "");
		xdrNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

		std::ignore = m_ws.relationships(); // create sheet relationships if not existing

		// ===== Trigger parentDoc to create drawing XML file and return it
		uint16_t sheetXmlNo = index();
		m_drawing =m_doc1->sheetDrawing(sheetXmlNo); // fetch drawing for this worksheet
		if (!m_drawing.valid())
			throw XLException("XLWorksheet::drawing(): could not create drawing XML");
		std::string drawingRelativePath = getPathARelativeToPathB(m_drawing.getXmlPath(), getXmlPath());
		XLRelationshipItem drawingRelationship;
		if (!m_ws.relationships().targetExists(drawingRelativePath))
			drawingRelationship = m_ws.relationships().addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
		else
			drawingRelationship = m_ws.relationships().relationshipByTarget(drawingRelativePath);
		if (drawingRelationship.empty())
			throw XLException("XLWorksheet::drawing(): could not add determine sheet relationship for Drawing");
		if (docElement.child("drawing").empty()) {
			XMLNode drawing = appendAndGetNode(docElement, "drawing", m_nodeOrder);
			if (drawing.empty())
				throw XLException("XLWorksheet::drawing(): could not add <drawing> element to worksheet XML");
			appendAndSetAttribute(drawing, "r:id", drawingRelationship.id());
		}

	}
	return m_drawing;
}
#endif
//-------------------class XLCell1---------------------------------------------------------

XLCell1::XLCell1() {}
XLCell1::XLCell1(XLDocument1 *doc1,XLWorksheet1 ws1,const XLCell c)
{
	m_doc1 = doc1;
	m_ws1 = ws1;
	m_c = c;
}

XLCell1::~XLCell1()
{
}

XLFont1 XLCell1::font()
{
	return XLFont1(m_doc1,*this);
}

XLBorders1 XLCell1::borders()
{
	return XLBorders1(m_doc1,*this);
}

XLCharacters1 XLCell1::characters(int16_t start, int16_t len)
{
	return XLCharacters1(m_doc1, *this, start, len);
}

XLBorder1 XLCell1::borders(int32_t index)
{
	return XLBorders1(m_doc1,*this).item(index);
}

XLCellValueProxy& XLCell1::value() {
	return m_c.value();
}

int32_t XLCell1::horizontalAlignment()
{
	return m_doc1->getintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL);
}

void XLCell1::setHorizontalAlignment(int32_t value)
{
	auto index = m_doc1->setintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL, value);
	m_c.setCellFormat(index);
}

void XLCell1::setHorizontalAlignment(std::string value)
{
	auto index = m_doc1->setcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_HORIZONTAL, value);
	m_c.setCellFormat(index);
}
int32_t XLCell1::verticalAlignment()
{
	return m_doc1->getintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL);
}

void XLCell1::setVerticalAlignment(int32_t value)
{
	auto index = m_doc1->setintstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL, value);
	m_c.setCellFormat(index);
}

void XLCell1::setVerticalAlignment(std::string value)
{
	auto index = m_doc1->setcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_VERTICAL, value);
	m_c.setCellFormat(index);
}

bool XLCell1::wraptext()
{
	return m_doc1->getboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_WRAPTEXT);
}

void XLCell1::setWraptext(bool value)
{
	auto index = m_doc1->setboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_WRAPTEXT, value);
	m_c.setCellFormat(index);
}

bool XLCell1::shrinktofit()
{
	return m_doc1->getboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_SHRINKTOFIT);
}

void XLCell1::setShrinktofit(bool value)
{
	auto index = m_doc1->setboolstyle(m_c.cellFormat(), MY_XLCELLFORMAT_ALIGNMENT, MY_ALIGN_SHRINKTOFIT, value);
	m_c.setCellFormat(index);
}
char * XLCell1::numberFormat()
{
	return m_doc1->getcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_NUMBERFORMATID, MY_NUMBERFORMAT_CODE);
}

void XLCell1::setNumberFormat(std::string value)
{
	auto index = m_doc1->setcharstyle(m_c.cellFormat(), MY_XLCELLFORMAT_NUMBERFORMATID, MY_NUMBERFORMAT_CODE, value);
	m_c.setCellFormat(index);
}

//--------------------class XLCharacters1--------------------------------------------------------------
XLCharacters1::XLCharacters1() {}

XLCharacters1::XLCharacters1(XLDocument1* doc1, XLCell1 c1, int16_t start, int16_t len)
{
	m_doc1 = doc1;
	m_c1 = c1;
	m_start = start;
	m_len = len;
}

XLCharacters1::~XLCharacters1()
{
}

XLFont1 XLCharacters1::font()
{
	return XLFont1(m_doc1,*this);
}

//------------------class XLCellRange1--------------------------------------------------
XLCellRange1::XLCellRange1() {}

XLCellRange1::XLCellRange1(XLDocument1* doc1, XLWorksheet1 ws1, const XLCellRange cr)
{
	m_doc1 = doc1;
	m_cr = cr;
	m_ws1 = ws1;
}

XLCellRange1::~XLCellRange1()
{
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

char* XLCellRange1::address()
{
	return (char *)m_cr.address().c_str();
}

XLBorder1 XLCellRange1::borders(int32_t index)
{
	return XLBorders1(m_doc1,*this).item(index);
}

XLFont1 XLCellRange1::font()
{
	return XLFont1(m_doc1,*this);
}
void XLCellRange1::merge()
{
	m_ws1.merge(m_cr.address());
}

void XLCellRange1::setpropchar(int32_t type, int32_t prop, std::string value)
{
	int32_t index;
	for (auto it = m_cr.begin(); it != m_cr.end(); ++it) {
		index = m_doc1->setcharstyle(it->cellFormat(), type, prop, value);
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
XLBorders1::XLBorders1() {}

XLBorders1::XLBorders1(XLDocument1* doc1, XLCell1 c1)
{
	m_t = 0;
	m_doc1 = doc1;
	m_c1 = c1;
}


XLBorders1::XLBorders1(XLDocument1 *doc1,XLCellRange1 cr1)
{
	m_t = 1;
	m_doc1 = doc1;
	m_cr1 = cr1;
}

XLBorders1::~XLBorders1()
{
}

XLBorder1 XLBorders1::item(int32_t index)
{
	return XLBorder1(m_doc1,*this,index);
}

//-----------------class XLBorder1--------------------------------------------------

XLBorder1::XLBorder1(XLDocument1 *doc1,XLBorders1 bs1,int32_t index)
{
	m_bs1 = bs1;
	m_doc1 = doc1;
	m_index = index;
}

XLBorder1::~XLBorder1()
{

}

int32_t XLBorder1::lineStyle()
{
	XLStyleIndex cf;
	if (m_bs1.t()==0) {
		XLCell c = m_bs1.c1().c();
		cf = c.cellFormat();
		return m_doc1->getintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index);
	}
	if (m_bs1.t()==1) {
		XLRECT rect;
//		XLCellRange1 cr1=m_bs1.cr1();
		m_bs1.cr1().rect(&rect);
		return m_bs1.cr1().ws1().cell(rect.top, (int16_t)rect.left).borders(0).lineStyle();
	}
}

void XLBorder1::setLineStyle(int32_t ls)
{	
	XLStyleIndex cf;
	if (m_bs1.t()==0) {
		XLCell c = m_bs1.c1().c();
		cf = c.cellFormat();
		cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, m_index, ls);
		c.setCellFormat(cf);
		return;
	}
	if (m_bs1.t()==1) {
		XLRECT rect;
		XLCellRange1 cr1 = m_bs1.cr1();
		XLWorksheet1 s1=cr1.ws1();
		cr1.rect(&rect);
		switch (m_index) {
		case 0:
			for (auto i = rect.top; i <= rect.bottom; i++) {
				XLCell1 c1=s1.cell(i,rect.left);
				XLCell c = c1.c();
				cf = c.cellFormat();
				cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, 0, ls);
				c.setCellFormat(cf);
			}
			break;
		case 1:
			for (auto i = rect.top; i <= rect.bottom; i++) {
				XLCell1 c1 = s1.cell(i,rect.right);
				XLCell c = c1.c();
				cf = c.cellFormat();
				cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, 1, ls);
				c.setCellFormat(cf);
			}
			break;
		case 2:
			for (auto i = rect.left; i <= rect.right; i++) {
				XLCell1 c1 = s1.cell(rect.top,i);
				XLCell c = c1.c();
				cf = c.cellFormat();
				cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, 2, ls);
				c.setCellFormat(cf);
			}
			break;
		case 3:
			for (auto i = rect.left; i <= rect.right; i++) {
				XLCell1 c1 = s1.cell(rect.bottom, i);
				XLCell c = c1.c();
				cf = c.cellFormat();
				cf = m_doc1->setintstyle(cf, MY_XLCELLFORMAT_BORDERINDEX, 3, ls);
				c.setCellFormat(cf);
			}
			break;
		}
	}
}

//------------------class XLFont1-----------------------------------------------------

XLFont1::XLFont1(XLDocument1 *doc1,XLCell1 c1)
{
	m_t = 0;
	m_doc1 = doc1;
	m_c1 = c1;
}

XLFont1::XLFont1(XLDocument1* doc1, XLCellRange1 cr1)
{
	m_t = 1;
	m_doc1 = doc1;
	m_cr1 = cr1;
}

XLFont1::XLFont1(XLDocument1* doc1, XLCharacters1 ch1)
{
	m_t = 2;
	m_doc1 = doc1;
	m_ch1 = ch1;
}

XLFont1::~XLFont1()
{
}

void XLFont1::setpropchar(int32_t type,int32_t prop, std::string value)
{
	XLStyleIndex index;
	if (m_t == 0) {
		XLCell c = m_c1.c();
		index = m_doc1->setcharstyle(c.cellFormat(), type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1.cr().begin(); it != m_cr1.cr().end(); ++it) {
			index = m_doc1->setcharstyle(it->cellFormat(), type, prop, value);
			it->setCellFormat(index);
		}
	}
	if (m_t == 2) {
		int32_t index, indexf;
		XLCHARACTERSTRUCT pp;
		XLCharacters1 ch1 = m_ch1;
		XLCell1 c1 = ch1.c1();
		XLWorksheet1 ws1 = c1.ws1();
		XLCell c = c1.c();
		pp.sheetno = ws1.index();
		pp.row = c.cellReference().row();
		pp.col = c.cellReference().column();
		pp.start = ch1.start();
		pp.len = ch1.len();
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
		XLCell c = m_c1.c();
		index = c.cellFormat();
		index = m_doc1->setintstyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1.cr().begin(); it != m_cr1.cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setintstyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}
	if (m_t == 2) {
		int32_t index, indexf;
		XLCHARACTERSTRUCT pp;
		XLCharacters1 ch1 = m_ch1;
		XLCell1 c1 = ch1.c1();
		XLWorksheet1 ws1 = c1.ws1();
		XLCell c = c1.c();
		pp.sheetno = ws1.index();
		pp.row = c.cellReference().row();
		pp.col = c.cellReference().column();
		pp.start = ch1.start();
		pp.len = ch1.len();
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
		XLCell c = m_c1.c();
		index = c.cellFormat();
		index = m_doc1->setboolstyle(index, type, prop, value);
		c.setCellFormat(index);
		return;
	}
	if (m_t == 1) {
		for (auto it = m_cr1.cr().begin(); it != m_cr1.cr().end(); ++it) {
			index = it->cellFormat();
			index = m_doc1->setboolstyle(index, type, prop, value);
			it->setCellFormat(index);
		}
		return;
	}
	if (m_t == 2) {
		int32_t index, indexf;
		XLCHARACTERSTRUCT pp;
		XLCharacters1 ch1 = m_ch1;
		XLCell1 c1 = ch1.c1();
		XLWorksheet1 ws1 = c1.ws1();
		XLCell c = c1.c();
		pp.sheetno =ws1.index();
		pp.row = c.cellReference().row();
		pp.col = c.cellReference().column();
		pp.start = ch1.start();
		pp.len = ch1.len();
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
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getcharstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_NAME);
}

void XLFont1::setName(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FONTINDEX,MY_XLFONT_NAME,value);
}


bool XLFont1::bold() {
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getboolstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_BOLD);
}

void XLFont1::setBold(bool value)
{
	setpropbool(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_NAME, value);
}

bool XLFont1::italic() {
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getboolstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_ITALIC);
}

void XLFont1::setItalic(bool value)
{
	setpropbool(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_ITALIC, value);
}

bool XLFont1::strikethrough() {
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getboolstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_STRIKETHROUGH);
}

void XLFont1::setStrikethrough(bool value)
{
	setpropbool(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_STRIKETHROUGH, value);
}

int XLFont1::underline() {
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_UNDERLINE);
}

void XLFont1::setUnderline(int value)
{
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_UNDERLINE, value);
}

int XLFont1::size() {
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	return m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_SIZE);
}

void XLFont1::setSize(int value)
{
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_SIZE, value);
}

bool XLFont1::superscript() {
	int n;
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	n=m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN);
	if (n == 2)return true;
	return false;
}

void XLFont1::setSuperscript(bool value)
{
	int32_t cfindex = 0; int n;
	if (value)n = 2;
	else n = 0;
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN,n);
}
bool XLFont1::subscript() {
	int n;
	XLCell c = m_c1.c();
	XLStyleIndex index = c.cellFormat();
	n = m_doc1->getintstyle(index, MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN);
	if (n == 1)return true;
	return false;
}

void XLFont1::setSubscript(bool value)
{
	int n;
	if (value)n = 1;
	else n = 0;
	setpropint(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_VERTALIGN, n);
}

void XLFont1::setColor(std::string value)
{
	setpropchar(MY_XLCELLFORMAT_FONTINDEX, MY_XLFONT_COLOR, value);
}


#ifdef MY_DRAWING
// ========== XLDrawing Member Functions

XLDrawing1::XLDrawing1(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
	if (xmlData->getXmlType() != XLContentType::Drawing)
		throw XLInternalError("XLDrawing constructor: Invalid XML data.");
	XMLDocument& doc = xmlDocument();
	if (doc.document_element().empty()) {  // handle a bad (no document element) drawing XML file
		std::string s1 = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		std::string s2 = "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"</xdr:wsDr>";
		doc.load_string((s1 + "\n" + s2).data(), pugi_parse_settings);
	}
	XMLNode rootNode = doc.document_element();
	XMLNode node = rootNode.first_child_of_type(pugi::node_element);

	while (not node.empty() && node.raw_name() == ShapeNodeNameDr) {
		XMLNode nextNode = node.next_sibling_of_type(pugi::node_element); // determine next node early because node may be invalidated by moveNode
		++m_shapeCount;
		node = nextNode;
	}

}

XMLNode XLDrawing1::rootNode() const {
	return xmlDocument().document_element();
}

XMLNode XLDrawing1::firstShapeNode() const
{
	XMLNode node = xmlDocument().document_element().first_child_of_type(pugi::node_element);
	while (not node.empty() && node.raw_name() != ShapeNodeNameDr)   // skip non shape nodes
		node = node.next_sibling_of_type(pugi::node_element);
	return node;
}

XMLNode XLDrawing1::lastShapeNode() const
{
	XMLNode node = xmlDocument().document_element().last_child_of_type(pugi::node_element);
	while (not node.empty() && node.raw_name() != ShapeNodeNameDr)
		node = node.previous_sibling_of_type(pugi::node_element);
	return node;

}

XMLNode XLDrawing1::shapeNode(uint32_t index) const
{
	using namespace std::literals::string_literals;

	XMLNode node{}; // scope declaration, ensures node.empty() when index >= m_shapeCount
	if (index < shapeCount()) {
		uint16_t i = 0;
		node = firstShapeNode();
		while (i != index && node.raw_name() == ShapeNodeNameDr) {
			++i;
			node = node.next_sibling_of_type(pugi::node_element);
		}
	}
	if (node.empty() || node.raw_name() != ShapeNodeNameDr)
		throw XLException("XLDrawing: shape index "s + std::to_string(index) + " is out of bounds"s);

	return node;
}

XMLNode XLDrawing1::shapeNode(std::string const& cellRef) const
{
	XLCellReference destRef(cellRef);
	uint32_t destRow = destRef.row() - 1;    // for accessing a shape: x:Row and x:Column are zero-indexed
	uint16_t destCol = destRef.column() - 1; // ..

	XMLNode node = firstShapeNode();
	while (not node.empty() && node.raw_name() == ShapeNodeNameDr) {
		if ((destRow == node.child("x:ClientData").child("x:Row").text().as_uint())
			&& (destCol == node.child("x:ClientData").child("x:Column").text().as_uint()))
			break; // found shape for cellRef

		do { // locate next shape node
			node = node.next_sibling_of_type(pugi::node_element);
		} while (not node.empty() && node.name() != ShapeNodeNameDr);
	}
	return node;
}

uint32_t XLDrawing1::shapeCount() const { return m_shapeCount; }

XLShape XLDrawing1::createShape()
{
	XMLNode rootNode = xmlDocument().document_element();
	XMLNode node = rootNode.append_child(ShapeNodeNameDr.c_str());
	m_shapeCount++;
	return XLShape(node);
}

XLShape XLDrawing1::shape(uint32_t index) const { return XLShape(shapeNode(index)); }

std::string XLDrawing1::data() const
{
	return xmlData();
}

XLDocument1* XLDrawing1::doc1()
{
	return m_doc1;
}
#endif

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

