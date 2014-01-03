package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook

class CellStyleEnhancer {
	private static def workbookMap = Collections.synchronizedMap([:])
	
	static void setWorkbook(CellStyle self, Workbook workbook) {
		workbookMap[System.identityHashCode(self) + ":" + self.index] = workbook
	}
	
	static Workbook getWorkbook(CellStyle self) {
		workbookMap[System.identityHashCode(self) + ":" + self.index]
	}
	
	static void setDataFormatString(CellStyle self, String formatString) {
		DataFormat fmt = self.workbook.creationHelper.createDataFormat()
		self.setDataFormat(fmt.getFormat(formatString))
	}
	
	static CellStyle combine(CellStyle self, CellStyle... others) {
		CellStyle combined = combineWithoutEnhancements(self, others)
		combined.workbook = self.workbook
		combined
	}
	
	private static CellStyle combineWithoutEnhancements(CellStyle self, CellStyle[] others) {
		Workbook workbook = self.workbook
		CellStyle base = workbook.getCellStyleAt(0 as short)
		
		def combined = [:]
		def attributes = ['alignment',
			'borderBottom', 'borderLeft', 'borderRight', 'borderTop',
			'bottomBorderColor', 'dataFormatString',
			'fillForegroundColor', 'fillBackgroundColor', // Note: Do FG before BG
			'fillPattern', 'hidden', 'indention', 'leftBorderColor',
			'locked', 'rightBorderColor', 'rotation', // 'shrinkToFit' - not present in XSSFCellStyle?
			'topBorderColor', 'verticalAlignment', 'wrapText']
		attributes.each {
			combined[it] = self[it]
		}
		
		others.each { other ->
			attributes.each { attr ->
				if (other[attr] != base[attr]) {
					combined[attr] = other[attr]
				}
			}
		}
			
		CellStyle combinedStyle = workbook.createCellStyle()
		use (CellStyleEnhancer) {
			combinedStyle.workbook = workbook
			combined.each { attribute, value ->
				combinedStyle[attribute] = value
			}
		}
		
		Font selfFont = workbook.getFontAt(self.fontIndex)
		Font[] otherFonts = others.collect { workbook.getFontAt(it.fontIndex) }
		Font combinedFont
		use (FontEnhancer) {
			selfFont.workbook = workbook
			combinedFont = selfFont.combine(otherFonts)
			combinedStyle.font = combinedFont
		}
		
		combinedStyle
	}
	
	static CellStyle enhance(CellStyle style, Workbook workbook) {
		style.metaClass {
			getWorkbook = { ->
				workbook
			}
			
			setDataFormatString = { String str ->
				CellStyleEnhancer.setDataFormatString(delegate, str)
			}
			
			getDataFormatString = { ->
				delegate.getDataFormatString()
			}
			
			combine = { CellStyle... others ->
				enhance(CellStyleEnhancer.combineWithoutEnhancements(delegate, others), workbook)
			}
		}
		style
	}
}
