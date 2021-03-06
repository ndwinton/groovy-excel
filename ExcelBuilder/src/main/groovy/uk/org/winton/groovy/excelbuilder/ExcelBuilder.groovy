/**
 * Groovy builder implementation for Excel 2007 workbooks using Apache POI.
 * 
 * Copyright (c) 2013, 2014 Neil Winton (neil@winton.org.uk)
 * All rights reserved.
 * 
 * This is open source software. See the file LICENCE.md for details.
 * 
 * @author Neil Winton <neil@winton.org.uk>
 */

package uk.org.winton.groovy.excelbuilder

import groovy.lang.Closure;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellValue
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelBuilder extends BuilderSupport {

	final Workbook workbook
	final Map<String, Sheet> sheets = [:]
	final Map<String, Font> fonts = [:]
	final Map<String, CellStyle> styles = [:]
	final FormulaEvaluator evaluator
	final File templateFile
	
	private String nextSheetName = 'Sheet1'
	private Sheet currentSheet
	private int nextRowNum = 0
	private Row currentRow
	private int nextColNum = 0
	
	ExcelBuilder() {
		this((File)null)
	}

	ExcelBuilder(String templateFileName) {
		this(new File(templateFileName))
	}
	
	ExcelBuilder(File templateFile) {
		this.templateFile = templateFile 
		if (templateFile != null) {
			workbook = loadTemplateAndPopulateSheetNames(templateFile)
		}
		else {
			workbook = new XSSFWorkbook()
		}
		evaluator = workbook.creationHelper.createFormulaEvaluator()
		createDefaultStyles()
	}


	private Workbook loadTemplateAndPopulateSheetNames(templateFile) {
		Workbook templateWorkbook = WorkbookFactory.create(templateFile)
		for (int sheetIndex = 0; sheetIndex < templateWorkbook.numberOfSheets; sheetIndex++) {
			sheets[templateWorkbook.getSheetName(sheetIndex)] = templateWorkbook.getSheetAt(sheetIndex)
		}
		return templateWorkbook
	}
	
	private createDefaultStyles() {
		createStyle('iso-date').dataFormatString = 'yyyy/mm/dd'
		createStyle('iso-datetime').dataFormatString = 'yyyy/mm/dd hh:mm:ss'
		createStyle('euro-date').dataFormatString = 'dd/mm/yyyy'
		createStyle('euro-datetime').dataFormatString = 'dd/mm/yyyy hh:mm:ss'
		createStyle('us-date').dataFormatString = 'mm/dd/yyyy'
		createStyle('us-datetime').dataFormatString = 'mm/dd/yyyy hh:mm:ss'
		styles['default-date'] = styles['iso-datetime']
		CellStyle base = CellStyleEnhancer.enhance(workbook.getCellStyleAt(0 as short), workbook)
		styles['default-numeric'] = styles['default-text'] = styles['default-boolean'] = base
	}
	
	@Override
	protected void setParent(Object parent, Object child) {
		// println "setParent($parent, $child)"
		if (child instanceof Sheet && !(parent instanceof ExcelBuilder)) {
			throw new IllegalArgumentException("sheets can only be created at the top level within a builder")
		}
	}

	@Override
	protected Object createNode(Object name) {
		// println "createNode($name)"
		switch (name) {
			case 'call':
				return this
				
			case 'sheet':
				return createSheet([name: uniqueSheetName()])
			
			case 'row':
				return createRow([:])
		}
		throw new IllegalArgumentException("Unknown builder operation: " + name + "()")
	}

	@Override
	protected Object createNode(Object name, Object value) {
		switch (name) {
			case 'call':
				return this
				
			case 'sheet':
				return createSheet([name: value])
			
			case 'row':
				return createRow([cells: value])
			
			case 'cell':
				return createCell([value: value])
				
			case 'font':
				return createFont(value)

			case 'style':
				return createStyle(value)
		}
		throw new IllegalArgumentException("Unknown builder operation: " + name + "(value)")
	}

	@Override
	protected Object createNode(Object name, Map attributes) {
		switch (name) {
			case 'call':
				return this
				
			case 'sheet':
				if (attributes.name == null) {
					attributes.name = nextSheetName++
				}
				return createSheet(attributes)
			
			case 'row':
				return createRow(attributes)
				
			case 'cell':
				return createCell(attributes)
		}
		throw new IllegalArgumentException("Unknown builder operation: " + name + "(attributes...)")
	}

	@Override
	protected Object createNode(Object name, Map attributes, Object value) {
		switch (name) {
			case 'call':
				return this
				
			case 'sheet':
				attributes.name = value
				return createSheet(attributes)
				
			case 'row':
				attributes.cells = value
				return createRow(attributes)
				
			case 'cell':
				attributes.value = value
				return createCell(attributes)
		}
		throw new IllegalArgumentException("Unknown builder operation: " + name + "(value, attributes...)")
	}

	@Override
	protected void nodeCompleted(Object parent, Object node) {
		// println("nodeCompleted($parent, $node)")
	}
	
	@Override
	protected void setClosureDelegate(Closure closure, Object node) {
		switch (node) {
			case Font:
			case CellStyle:
				closure.setDelegate(node)
				break
			
			default:
				closure.setDelegate(this)
				break
		}
	}
	
	private Sheet createSheet(Map attributes) {
		
		def name = attributes.name
		currentSheet = workbook.getSheet(name) ?: workbook.createSheet(name)
		sheets[name] = currentSheet
		SheetEnhancer.enhance(currentSheet)
		
		currentSheet.active = attributes.active ?: false
		currentSheet.hidden = attributes.hidden ?: false
		if (attributes.width >= 0) {
			currentSheet.defaultColumnWidthInChars = attributes.width
		}
		if (attributes.height >= 0) {
			currentSheet.defaultRowHeightInPoints = attributes.height
		}
		return currentSheet
	}
	
	private String uniqueSheetName() {
		while (workbook.getSheet(nextSheetName) != null) {
			nextSheetName++
		}
		nextSheetName
	}
	
	private Row createRow(Map attributes) {
		if (currentSheet == null) {
			throw new IllegalArgumentException("row can't be created without a previously defined sheet")
		}
		nextColNum = 0
		if (attributes.row != null) {
			nextRowNum = attributes.row
		}
		currentRow = findOrCreateRow(nextRowNum++)
		applyStyles(currentRow, attributes.style)
		
		if (attributes.height != null) {
			currentRow.heightInPoints = attributes.height
		}
		
		currentRow.metaClass.getAttributes = { -> attributes }
		
		attributes.cells?.each { value ->
			createCell([value: value, style: attributes.style, width: attributes.width])
		}
		
		currentRow
	}

	private Row findOrCreateRow(rowNum) {
		currentSheet.getRow(rowNum) ?: currentSheet.createRow(rowNum)
	}
	
	private Cell createCell(Map attributes) {
		def value = attributes.value
		if (attributes.row != null) {
			createRow([row: attributes.row])
		}
		if (currentRow == null) {
			createRow([:])
		}
		nextColNum = attributes.column != null ? attributes.column : nextColNum
		Cell cell = currentRow.getCell(nextColNum++, Row.CREATE_NULL_AS_BLANK)
		
		switch (value) {
			case null:
				// Do not overwrite current value if null
				break
	
			case Number:
				cell.setCellValue(value)
				cell.cellStyle = styles['default-numeric']
				break
				
			case Boolean:
				cell.setCellValue(value)
				cell.cellStyle = styles['default-boolean']
				break
				
			case Date:
			case Calendar:
				cell.setCellValue(value)
				cell.cellStyle = styles['default-date']
				break
	
			case String:
			default:
				setFormulaOrStringCellValue(cell, value)
				break
		}
		
		applyStyles(cell, attributes.style ?: findStyle(currentRow.rowStyle))

		setRowHeight(currentRow, attributes.height, attributes.force)				
		setCellWidth(cell, attributes.width != null ? attributes.width : currentRow.attributes.width, attributes.force)
		
		cell
	}
	
	private setFormulaOrStringCellValue(Cell cell, String value) {
		def style = styles['default-text']
		if (value.startsWith("=")) {
			value = value[1..-1]
			cell.cellFormula = value
			evaluator.evaluateFormulaCell(cell)
			switch (cell.cachedFormulaResultType) {
				case Cell.CELL_TYPE_BOOLEAN:
					style = styles['default-boolean']
					break
					
				case Cell.CELL_TYPE_NUMERIC:
					style = styles['default-numeric']
					break
					
				case Cell.CELL_TYPE_STRING:
					style = styles['default-text']
					break
					
				default:
					break
			}
		}
		else if (value.startsWith("'")) {
			cell.setCellValue(value[1..-1])
		}
		else {
			cell.setCellValue(value)
		}
		cell.cellStyle = style
	}
	
	private void setRowHeight(row, height, force) {
		if (height == null) {
			return
		}
		
		if (height > row.heightInPoints || force) {
			row.heightInPoints = height
		}
	}
	
	private void setCellWidth(cell, width, force) {
		if (width == null) {
			return
		}
		
		if (width > currentSheet.getColumnWidthInChars(cell.columnIndex) || force) {
			currentSheet.setColumnWidthInChars(cell.columnIndex, width)
		}
	}
	
	private Font createFont(name) {
		fonts[name] = workbook.createFont()
		FontEnhancer.enhance(fonts[name], workbook)
		if (!styles[name]) {
			createStyle(name)
			styles[name].font = fonts[name]
		}
		fonts[name]
	}
	
	private CellStyle createStyle(name) {
		styles[name] = workbook.createCellStyle()
		CellStyleEnhancer.enhance(styles[name], workbook)
		styles[name]
	}
	
	private void applyStyles(entity, styleNameList) {
		if (styleNameList != null) {
            styleNameList = ensureStyleNameListIsAList(styleNameList)
			def name = styleNameList.join('+')
			if (styles[name] == null) {
				def styleList = []
				styleNameList.each {
					if (styles[it]) {
						styleList << styles[it]
					}
					else {
						throw new IllegalArgumentException("undefined style name: " + it)
					}
				}
				styles[name] = styleList[0].combine(*styleList)
			}
			switch (entity) {
				case Cell:
					entity.cellStyle = styles[name]
					break
				
				case Row:
					entity.rowStyle = styles[name]
					break
			}
		}
	}
	
	private def List ensureStyleNameListIsAList(styleNameList) {
		switch (styleNameList) {
		case String:
			return [styleNameList]

		case List:
		return styleNameList

		default:
			throw new IllegalArgumentException("style list must be a string or list")
		}
	}

	private def findStyle(style) {
		styles.findResult { k, v -> v == style ? k : null }
	}
}
