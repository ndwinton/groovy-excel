package uk.org.winton.groovy.excelbuilder

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelBuilder extends BuilderSupport {

	final Workbook workbook
	final Map<String, Sheet> sheets = [:]
	
	private String nextSheetName = 'Sheet1'
	private Sheet currentSheet
	private int nextRowNum = 0
	private Row currentRow
	private int nextColNum = 0
	
	ExcelBuilder() {
		workbook = new XSSFWorkbook()
	}

	@Override
	protected void setParent(Object parent, Object child) {
		println "setParent($parent, $child)"
		if (child instanceof Sheet && !(parent instanceof ExcelBuilder)) {
			throw new IllegalArgumentException("sheets can only be created at the top level within a builder")
		}
	}

	@Override
	protected Object createNode(Object name) {
		println "createNode($name)"
		switch (name) {
			case 'call':
				return this
				
			case 'sheet':
				return createSheet([name: nextSheetName++])
			
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
		}
		throw new IllegalArgumentException("Unknown builder operation: " + name + "(value)")
	}

	@Override
	protected Object createNode(Object name, Map attributes) {
		switch (name) {
			case 'call':
				return this
				
			case 'sheet':
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
		println("nodeCompleted($parent, $node)")
	}
	
	private Sheet createSheet(Map attributes) {
		
		def name = attributes.name
		currentSheet = workbook.createSheet(name)
		sheets[name] = currentSheet
		enrichCurrentSheetMetaClass()
		
		currentSheet.active = attributes.active ?: false
		currentSheet.hidden = attributes.hidden ?: false
		
		return currentSheet
	}
	
	/**
	 * Metaclass modification add:
	 * 
	 * hidden property (read/write)
	 * 		Sets the sheet hidden or shown (default)
	 * active property (read/write)
	 * 		Makes the sheet active (and others inactive)
	 * rows property (readonly)
	 * 		Returns iterator for the rows within the sheet.
	 * 		Note that the actual row number may be different from its
	 * 		position in the sequence.
	 */
	private void enrichCurrentSheetMetaClass() {
		def index = workbook.getSheetIndex(currentSheet)

		// NB: 'index' is available within this closure
		
		currentSheet.metaClass {
			setActive = { boolean on ->
				if (on) {
					workbook.setActiveSheet(index)
				}
			}
			getActive = { ->
				workbook.getActiveSheetIndex() == index
			}
			
			setHidden = { boolean hide ->
				workbook.setSheetHidden(index, hide)
				
			}
			getHidden = { ->
				workbook.isSheetHidden(index)
			}
		}
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
		attributes.cells?.each { value ->
			createCell([value: value])
		}
		currentRow
	}

	private Row findOrCreateRow(rowNum) {
		currentSheet.getRow(rowNum) ?: currentSheet.createRow(rowNum)
	}
	
	private Cell createCell(Map attributes) {
		def value = attributes.value ?: ''
		if (attributes.row != null) {
			createRow([row: attributes.row])
		}
		nextColNum = attributes.column != null ? attributes.column : nextColNum
		Cell cell = currentRow.getCell(nextColNum++, Row.CREATE_NULL_AS_BLANK)
		
		switch (value) {
			case null:
			case Number:
			case Boolean:
			case Date:
			case Calendar:
			case String:
				cell.setCellValue(value)
				break
			
			default:
				cell.setCellValue(value.toString())
				break
		}
		cell
	}
}
