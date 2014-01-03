/**
 * Groovy builder implementation for Excel 2007 workbooks using Apache POI.
 * 
 * <pre>
 * Copyright (c) 2013, Neil Winton (neil@winton.org.uk)
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without modification,
 * are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice, this
 * list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 * this list of conditions and the following disclaimer in the documentation and/or
 * other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
 * OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT
 * SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
 * SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT
 * OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
 * HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
 * TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE,
 * EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * </pre>
 * 
 * @author Neil Winton <neil@winton.org.uk>
 */

package uk.org.winton.groovy.excelbuilder

import groovy.lang.Closure;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelBuilder extends BuilderSupport {

	final Workbook workbook
	final Map<String, Sheet> sheets = [:]
	final Map<String, Font> fonts = [:]
	final Map<String, CellStyle> styles = [:]
	
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
		currentSheet = workbook.createSheet(name)
		sheets[name] = currentSheet
		SheetEnhancer.enhance(currentSheet)
		
		currentSheet.active = attributes.active ?: false
		currentSheet.hidden = attributes.hidden ?: false
		
		return currentSheet
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
			createCell([value: value, style: attributes.style])
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
		applyStyles(cell, attributes.style)
		cell
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
	
	private void applyStyles(Cell cell, styleNameList) {
		if (styleNameList != null) {
			if (!styleNameList.respondsTo('join')) {
				styleNameList = [styleNameList]
			}
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
			cell.setCellStyle(styles[name])
		}
	}
}
