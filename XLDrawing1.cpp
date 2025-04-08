#include "XLDrawing1.hpp"

XLDrawing1& XLDocument::sheetDrawing(uint16_t sheetXmlNo)
{
	using namespace std::literals::string_literals;
	std::string drawingFilename = "xl/drawings/drawing"s + std::to_string(sheetXmlNo) + ".xml"s;

	if (!m_archive.hasEntry(drawingFilename)) {
		// ===== Create the sheet drawing file within the archive
		m_archive.addEntry(drawingFilename, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");  // empty XML file, class constructor will do the rest
		if (!m_contentTypes.PartNameExists("/" + drawingFilename))
			m_contentTypes.addOverride("/" + drawingFilename, XLContentType::Drawing);                          // add content types entry
	}
	constexpr const bool DO_NOT_THROW = true;
	XLXmlData* xmlData = getXmlData((const std::string&)drawingFilename, DO_NOT_THROW);
	if (xmlData == nullptr) // if not yet managed: add the sheet drawing file to the managed files
		xmlData = &m_data.emplace_back(this, drawingFilename, "", XLContentType::Drawing);

	return XLDrawing1(xmlData);
}

bool XLDocument::hasSheetDrawing(uint16_t sheetXmlNo) const
{
	using namespace std::literals::string_literals;
	return m_archive.hasEntry("xl/drawings/drawing"s + std::to_string(sheetXmlNo) + ".xml"s);
}

bool XLWorksheet::hasDrawing() const
{
	return hasSheetDrawing(m_index);
}

XLDrawing1& XLWorksheet::drawing()
{
	if (!m_drawing.valid()) {
		// ===== Append xdr namespace attribute to worksheet if not present

		XMLNode docElement = xmlDocument().document_element();
		XMLAttribute xdrNamespace = appendAndGetAttribute(docElement, "xmlns:xdr", "");
		xdrNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

		std::ignore = m_ws.relationships(); // create sheet relationships if not existing

		// ===== Trigger parentDoc to create drawing XML file and return it
		uint16_t sheetXmlNo = index();
		m_drawing = sheetDrawing(sheetXmlNo); // fetch drawing for this worksheet
		if (!m_drawing.valid())
			throw XLException("XLWorksheet::drawing(): could not create drawing XML");
		std::string drawingRelativePath = getPathARelativeToPathB(m_drawing.getXmlPath(), getXmlPath());
		XLRelationshipItem drawingRelationship;
		if (!relationships().targetExists(drawingRelativePath))
			drawingRelationship = relationships().addRelationship(XLRelationshipType::Drawing, drawingRelativePath);
		else
			drawingRelationship = relationships().relationshipByTarget(drawingRelativePath);
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

bool XLWorksheet::hasDrawing()  const { return parentDoc().hasSheetDrawing(sheetXmlNumber());} 
