#include <OpenXLSX.hpp>

using namespace OpenXLSX;

#define MY_XLCELLFORMAT_NUMBERFORMATID 0
#define MY_XLCELLFORMAT_FONTINDEX 1
#define MY_XLCELLFORMAT_FILLINDEX 2
#define MY_XLCELLFORMAT_BORDERINDEX 3
#define MY_XLCELLFORMAT_ALIGNMENT 4
#define MY_XLCELLFORMAT_XFID 5

#define MY_XLFONT_NAME 0
#define MY_XLFONT_CHARSET 1
#define MY_XLFONT_FAMILY 2
#define MY_XLFONT_SIZE 3
#define MY_XLFONT_COLOR 4
#define MY_XLFONT_BOLD 5
#define MY_XLFONT_ITALIC 6
#define MY_XLFONT_STRIKETHROUGH  7
#define MY_XLFONT_UNDERLINE 8
#define MY_XLFONT_OUTLINE 9
#define MY_XLFONT_SHADOW 10
#define MY_XLFONT_CONDENSE 11
#define MY_XLFONT_EXTEND 12
#define MY_XLFONT_VERTALIGN 13
#define MY_XLFONT_SCHEME 14

#define MY_NUMBERFORMAT_ID 0
#define MY_NUMBERFORMAT_CODE 1

#define MY_BORDER_LEFT 0
#define MY_BORDER_RIGHT 1
#define MY_BORDER_TOP 2
#define MY_BORDER_BOTTOM 3
#define MY_BORDER_DIAGONALUP 4 
#define MY_BORDER_DIAGONALDOWN 5
#define MY_BORDER_VERTICAL 6
#define MY_BORDER_HORIZONTAL 7
#define MY_BORDER_OUTLINE 8
#define MY_BORDER_DIAGONAL 9

#define MY_FILL_FILLTYPE 0
#define MY_FILL_GRADIENTTYPE 1
#define MY_FILL_DEGREE 2
#define MY_FILL_LEFT 3
#define MY_FILL_RIGHT 4
#define MY_FILL_TOP 5
#define MY_FILL_BOTTOM 6
#define MY_FILL_PATTERNTYPE 7
#define MY_FILL_COLOR 8
#define MY_FILL_BACKGROUNDCOLOR 9

#define MY_ALIGN_HORIZONTAL 0
#define MY_ALIGN_VERTICAL 1
#define MY_ALIGN_WRAPTEXT 2
#define MY_ALIGN_JUSTIFYLASTLINE 3
#define MY_ALIGN_SHRINKTOFIT 4
#define MY_ALIGN_TEXTROTATION 5
#define MY_ALIGN_INDENT 6
#define MY_ALIGN_RELATIVEINDENT 7
#define MY_ALIGN_READINGORDER 8

#pragma pack(1)

typedef struct XLRECT {
	int32_t left;
	int32_t top;
	int32_t right;
	int32_t bottom;
} XLRECT;

typedef struct XLCOLORSTRUCT
{
	uint8_t alpha;
	uint8_t red;
	uint8_t green;
	uint8_t blue;
}XLCOLORSTRUCT;

typedef struct XLFONTSTRUCT
{
	char name[32];
	int32_t charset;
	int32_t family;
	int32_t size;
	union {
		XLCOLORSTRUCT color;
		int32_t rgb;
	}i;
	int8_t bold;
	int8_t italic;
	int8_t strikethrough;
	int8_t underline;
	int8_t outline;
	int8_t shadow;
	int8_t condense;
	int8_t extend;
	int8_t vertalign;
	int8_t scheme;
	int16_t unsave;
}XLFONTSTRUCT;

typedef struct XLALIGNSTRUCT
{
	int8_t horizontal;
	int8_t vertical;
	int8_t wraptext;
	int8_t justifylastline;
	int8_t shrinktofit;
	int16_t textrotation;
	int32_t indent;
	int32_t relativeindent;
	int32_t readingorder;
}XLALIGNSTRUCT;

typedef struct XLFILLSTRUCT
{
	int8_t filltype;
	int8_t gradienttype;
	double degree;
	double left;
	double right;
	double top;
	double bottom;
	int16_t patterntype;
	XLCOLORSTRUCT color;
	int32_t backgroundcolor;
	int16_t unsave;
}XLFILLSTRUCT;

typedef struct XLNUMBERFORMATSTRUCTEMBED
{
	int32_t id;
	char formatcode[60];
}XLNUMBERFORMATSTRUCTEMBED;

typedef struct XLNUMBERFORMATSTRUCT
{
	int32_t id;
	char formatcode[60];
	int16_t unsave;
}XLNUMBERFORMATSTRUCT;

typedef struct XLDATABARCOLORSTRUCT
{
	XLCOLORSTRUCT rgb;
	double tint;
	uint32_t indexed;
	uint32_t theme;
	int8_t automatic;

}XLDATABARCOLORSTRUCT;


typedef struct XLLINESTRUCT
{
	uint8_t style;
	XLDATABARCOLORSTRUCT color;
}XLLINESTRUCT;

typedef struct XLBORDERSTRUCT
{
	int8_t outline;
	int8_t diagonalup;
	int8_t diagonaldown;
	int8_t unsave;
	XLLINESTRUCT left;
	XLLINESTRUCT right;
	XLLINESTRUCT top;
	XLLINESTRUCT bottom;
	XLLINESTRUCT vertical;
	XLLINESTRUCT horizontal;
	XLLINESTRUCT diagonal;

}XLBORDERSTRUCT;


typedef struct XLCELLFORMATSTRUCT
{
	uint32_t numberformatid;
	uint32_t fontindex;
	uint32_t fillindex;
	uint32_t borderindex;
	uint32_t xfid;
	XLALIGNSTRUCT alignment;
	uint32_t unsave;
}XLCELLFORMATSTRUCT;

typedef struct XLCHARACTERSTRUCT
{
	int16_t sheetno;
	int32_t row;
	int16_t col;
	int16_t start;
	int16_t len;
	int32_t indexf;
}XLCHARACTERSTRUCT;

#pragma pack()

class XLWorkbook1;
class XLDrawing1;
class XLWorksheet1;
class XLCell1;
class XLCellRange1;
class XLCharacters1;
class XLFont1;
class XLBorders1;
class XLBorder1;
class XLFill1;

class XLDocument1
{
	friend class XLWorkbook1;
	friend class XLFont1;
public:
	XLDocument1();
	~XLDocument1();

	void getallstyles();
	void setallstyles();

	bool getboolstyle(int32_t index, int32_t type, int32_t prop);
	int32_t getintstyle(int32_t index, int32_t type, int32_t prop);
	char* getcharstyle(int32_t index, int32_t type, int32_t prop);

	char* findnumberformat(int id);
	int32_t findnumberformat(char* code);
	int32_t getnumberformatnextfreeid();
	int32_t createnumberformat(char* code);

	int32_t findfont(void* p);
	int32_t createfont(void* p);

	int32_t findcharacter(void* p);
	int32_t createcharacter(void* p);

	int32_t findfill(void* p);
	int32_t createfill(void* p);

	int32_t findcellformat(XLCELLFORMATSTRUCT* p);
	int32_t countcellformat(int32_t type, int32_t n);
	int32_t createcellformat(void* p);

	int32_t createborder(void* p);
	int32_t findborder(void* p);

	int32_t setboolstyle(int32_t index, int32_t type, int32_t prop, bool value);
	int32_t setintstyle(int32_t index, int32_t type, int32_t prop, int value);
	int32_t setcharstyle(int32_t index, int32_t type, int32_t prop, std::string value);
	
	XLDocument* doc();
	void create(const std::string& fileName, bool forceOverwrite);
	void open(const std::string& fileName);
	void save();
	void close();
	XLWorkbook1 workbook();
	void copyRange(int sheetXmlNo, XLRECT* from, XLRECT* to);
#ifdef MY_DRAWING
	char* shapeAttribute(int sheetXmlNo, int shapeNo, char* path);
	void setShapeAttribute(int sheetXmlNo, int shapeNo, char* path, char* attribute, char* value);
	int appendPictures(int sheetXmlNo, void* buffer, int bufferlen, char* ext, XLRECT* rect);
	bool hasSheetDrawing(uint16_t sheetXmlNo) const;
	XLDrawing1& sheetDrawing(uint16_t sheetXmlNo);
#endif
private:
	XLDocument* m_doc;
	int m_save;
	int m_begin;

	XLFONTSTRUCT* m_fonts = NULL;
	int m_fontcount = 0;

	XLFILLSTRUCT* m_fills = NULL;
	int m_fillcount = 0;

	XLCELLFORMATSTRUCT* m_cellformat = NULL;
	int m_cellformatcount = 0;

	int m_numberformatnextfreeid = 165;
	XLNUMBERFORMATSTRUCT* m_numberformat = NULL;
	int m_numberformatcount = 0;

	XLBORDERSTRUCT* m_borders = NULL;
	int m_bordercount = 0;

	XLCHARACTERSTRUCT* m_characters = NULL;
	int m_charactercount = 0;
};

class XLWorkbook1 : public XLWorkbook
{
friend XLDocument1;
public :
	XLWorkbook1(XLDocument1 *doc1, const XLWorkbook wb);
	~XLWorkbook1();
	XLWorksheet1 worksheet(int16_t n);
	XLWorksheet1 worksheet(std::string name);
private :
	XLDocument1* m_doc1;
	XLWorkbook m_wb;
};

class XLWorksheet1 : public XLXmlFile
{
friend XLDocument1;
friend XLWorksheet;
public:
	XLWorksheet1();
	XLWorksheet1(XLDocument1* doc1, XLWorksheet ws);
	~XLWorksheet1();
	XLWorksheet ws() { return m_ws; };
	XLCell1 cell(const std::string &address);
	XLCell1 cell(int32_t row, int16_t column);
	XLCellRange1 range(const std::string &address);
	void merge(const std::string &address);
	XLColumn column(int16_t column);
	XLRow row(int32_t row);
	void setSelected(bool sel);
	int16_t index() { return m_index; };
#ifdef MY_DRAWING
	bool hasDrawing() const;
	XLDrawing1& drawing();
#endif
private:
	XLDocument1* m_doc1;
	XLWorksheet m_ws;
	int16_t m_index;
#ifdef MY_DRAWING
	XLDrawing1& m_drawing = drawing();
#endif
	inline static const std::vector< std::string_view > m_nodeOrder = {
		   "formula",
		   "colorScale",
		   "dataBar",
		   "iconSet",
		   "extLst"
	};
};

class XLCell1 : public XLCell
{
friend XLDocument1;
public :
	XLCell1();
	XLCell1(XLDocument1* doc1, XLWorksheet1 ws1, const XLCell c);
	~XLCell1();
	XLCellValueProxy& value();
	XLFont1 font();
	XLBorders1 borders();
	XLBorder1 borders(int32_t index);
	XLCharacters1 characters(int16_t start, int16_t len);
	XLDocument1* doc1() { return m_doc1; };
	XLWorksheet1 ws1() { return m_ws1; };
	const XLCell c() { return m_c; };
	int32_t horizontalAlignment();
	void setHorizontalAlignment(int32_t value);
	void setHorizontalAlignment(std::string value);
	int32_t verticalAlignment();
	void setVerticalAlignment(int32_t value);
	void setVerticalAlignment(std::string value);
	bool wraptext();
	void setWraptext(bool value);
	bool shrinktofit();
	void setShrinktofit(bool value);
	char* numberFormat();
	void setNumberFormat(std::string value);
private:
	XLDocument1* m_doc1;
	XLWorksheet1 m_ws1 = XLWorksheet1();
	XLCell m_c;
};

class XLCellRange1 : public XLCellRange
{
friend XLDocument1;
public:
	XLCellRange1();
	XLCellRange1(XLDocument1* doc1, XLWorksheet1 ws1, const XLCellRange cr);
	~XLCellRange1();
	XLDocument1* doc1() { return m_doc1; };
	XLWorksheet1 ws1() { return m_ws1; };
	const XLCellRange cr() { return m_cr; };
	void rect(XLRECT *rect);
	XLFont1 font();
	void merge();
	char * address();
	XLBorder1 borders(int32_t index);
	void setpropchar(int32_t type, int32_t prop, std::string value);
	void setpropint(int32_t type, int32_t prop, int32_t value);
	void setpropbool(int32_t type, int32_t prop, bool value);
	void setHorizontalAlignment(int32_t value);
	void setHorizontalAlignment(std::string value);
	void setVerticalAlignment(int32_t value);
	void setVerticalAlignment(std::string value);
	void setWraptext(bool value);
	void setShrinktofit(bool value);
	void setNumberFormat(std::string value);
private:
	XLDocument1* m_doc1;
	XLCellRange m_cr;
	XLWorksheet1 m_ws1=XLWorksheet1();
};

class XLCharacters1
{
public:
	XLCharacters1();
	XLCharacters1(XLDocument1 *doc1,XLCell1 c1,int16_t start,int16_t len);
	~XLCharacters1();
	XLDocument1* doc1() { return m_doc1; };
	XLCell1 c1() { return m_c1; };
	int16_t start() { return m_start; };
	int16_t len() { return m_len; };
	XLFont1 font();
private:
	XLDocument1 *m_doc1;
	XLCell1 m_c1=XLCell1();
	int16_t m_start;
	int16_t m_len;
};

class XLBorders1
{
public:
	XLBorders1();
	XLBorders1(XLDocument1 *doc1,XLCell1 c1);
	XLBorders1(XLDocument1 *doc1,XLCellRange1 cr1);
	~XLBorders1();
	XLDocument1* doc1() { return m_doc1; };
	int32_t t() { return m_t; };
	XLCell1 c1() { return m_c1; };
	XLCellRange1 cr1() { return m_cr1; };
	XLBorder1 item(int32_t n);
private:
	int32_t m_t;
	XLDocument1* m_doc1;
	XLCell1 m_c1=XLCell1();
	XLCellRange1 m_cr1=XLCellRange1();
};

class XLBorder1
{
public:
	XLBorder1(XLDocument1 *doc1,XLBorders1 bs,int32_t index);
	~XLBorder1();
	void setLineStyle(int32_t ls);
	int32_t lineStyle();
private:
	XLDocument1* m_doc1;
	XLBorders1 m_bs1=XLBorders1();
	int32_t m_index;
};

class XLFill1
{
public:
	XLFill1();
	~XLFill1();
};

class XLFont1
{
	friend XLDocument1;
	friend XLCell1;
	friend XLCellRange1;
	friend XLCharacters1;
public:
	XLFont1(XLDocument1 *doc1,XLCell1 c1);
	XLFont1(XLDocument1* doc1, XLCellRange1 cr1);
	XLFont1(XLDocument1* doc1, XLCharacters1 ch1);
	~XLFont1();
	XLCell1 c1() { return m_c1; };
	XLCellRange1 cr1(){ return m_cr1; };
	XLCharacters1 ch1() { return m_ch1; };
	char * name();
	void setSize(int32_t value);
	int32_t size();
	void setName(std::string value);
	void setpropchar(int32_t type,int32_t prop, std::string value);
	void setpropint(int32_t type, int32_t prop, int32_t value);
	void setpropbool(int32_t type, int32_t prop, bool value);
	bool bold();
	void setBold(bool value);
	bool italic();
	void setItalic(bool value);
	bool strikethrough();
	void setStrikethrough(bool value);
	void setUnderline(int32_t value);
	int32_t underline();
	bool superscript();
	void setSuperscript(bool value);
	bool subscript();
	void setSubscript(bool value);
	void setColor(std::string value);
private :
	XLDocument1* m_doc1;
	int32_t m_t = 0;
	XLCell1 m_c1=XLCell1();
	XLCellRange1 m_cr1=XLCellRange1();
	XLCharacters1 m_ch1=XLCharacters1();
};

#ifdef MY_DRAWING
class XLDrawing1 : public XLXmlFile
{
	friend class XLWorksheet;   // for access to XLXmlFile::getXmlPath
	friend class XLWorksheet1;
public:
	XLDrawing1() : XLXmlFile(nullptr) {};
	XLDrawing1(XLXmlData* xmlData);
	XLDrawing1(const XLDrawing1& other) = default;
	XLDrawing1(XLDrawing1&& other) noexcept = default;
	~XLDrawing1() = default;

	XLDrawing1& operator=(const XLDrawing1&) = default;
	XLDrawing1& operator=(XLDrawing1&& other) noexcept = default;

	XMLNode shapeNode(std::string const& cellRef) const;
	XMLNode shapeNode(uint32_t index) const;

	XLShape shape(uint32_t index) const;
	XLShape createShape();

	uint32_t shapeCount() const;

	XMLNode rootNode() const;

	std::string data() const;
	XLDocument1* doc1();

private: 
	XMLNode firstShapeNode() const;
	XMLNode lastShapeNode() const;

private:
	uint32_t m_shapeCount{ 0 };
	uint32_t m_lastAssignedShapeId{ 0 };
	std::string m_defaultShapeTypeId{};
	XLDocument1* m_doc1;

};
#endif

/* Demo RTF - included
<r>
	<t>pri</t>
</r>
<r>
	<rPr>
		<u/>
		<i/>
		<b/>
	</rPr>
	<t>vet</t>
</r>
*/
