#ifndef OPENXLSX_XLDRAWING1_HPP
#define OPENXLSX_XLDRAWING1_HPP

// ===== External Includes ===== //
#include <cstdint>      // uint8_t, uint16_t, uint32_t
#include <ostream>      // std::basic_ostream

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"

using namespace OpenXLSX;

extern const std::string ShapeTypeNodeNameDr;

class XLDrawing1 : public XLXmlFile
{
	friend class XLWorksheet;
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
