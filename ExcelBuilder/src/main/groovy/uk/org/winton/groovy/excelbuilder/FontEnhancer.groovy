package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook

class FontEnhancer {
	private static final def workbookMap = [:]
	
	static void setWorkbook(Font self, Workbook workbook) {
		workbookMap[System.identityHashCode(self) + ":" + self.index] = workbook
	}
	
	static Workbook getWorkbook(Font self) {
		workbookMap[System.identityHashCode(self) + ":" + self.index]
	}
	
	// Not actually necessary for XSSFFont class, which already has a getBold/setBold,
	// but it still works and also works for HSSFFont
	static void setBold(Font self, boolean bold) {
		self.boldweight = bold ? Font.BOLDWEIGHT_BOLD : Font.BOLDWEIGHT_NORMAL
	}
	
	static boolean getBold(Font self) {
		self.boldweight == Font.BOLDWEIGHT_BOLD
	}
	
	static void setFontHeightInPoints(Font self, Number size) {
		self.fontHeight = Math.min((size * 20) as float, Short.MAX_VALUE as float) as short
	}
	
	static Number getFontHeightInPoints(Font self) {
		self.fontHeight / 20.0
	}
	
	static Font combine(Font self, Font... others) {
		Font combined = combineWithoutEnhancements(self, others)
		combined.workbook = self.workbook
		combined
	}
	
	private static Font combineWithoutEnhancements(Font self, Font[] others) {
		Workbook workbook = self.workbook
		Font base = workbook.getFontAt(0 as short)
		
		def combined = [:]
		def attributes = ['boldweight', 'charSet', 'color', 'fontHeight',
				'fontName', 'italic', 'strikeout', 'typeOffset', 'underline']
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
			
		Font combinedFont = workbook.findFont(combined.boldweight, combined.color, combined.fontHeight,
			combined.fontName, combined.italic, combined.strikeout, combined.typeOffset, combined.underline)
		
		if (!combinedFont || combined.charSet != combinedFont.charSet) {
			combinedFont = workbook.createFont()
			combined.each { attribute, value ->
				combinedFont[attribute] = value
			}
		}
		combinedFont
	}
	
	static Font enhance(Font font, Workbook workbook) {
		font.metaClass {
			// If necessary, setWorkbook could be implemented as follows ...
			//
			// setWorkbook = { Workbook wb ->
			//	delegate.metaClass.getWorkbook = { -> wb }
			// }
			
			getWorkbook = { ->
				workbook
			}
			
			setBold = { boolean bold ->
				FontEnhancer.setBold(delegate, bold)
			}
			
			getBold = { ->
				FontEnhancer.getBold(delegate)
			}
			
			setFontHeightInPoints = { Number size ->
				FontEnhancer.setFontHeightInPoints(delegate, size)
			}
			
			getFontHeightInPoints = { ->
				FontEnhancer.getFontHeightInPoints(delegate)
			}
			
			combine = { Font... others ->
				enhance(FontEnhancer.combineWithoutEnhancements(delegate, others), workbook)
			}
		}
		font
	}
}
