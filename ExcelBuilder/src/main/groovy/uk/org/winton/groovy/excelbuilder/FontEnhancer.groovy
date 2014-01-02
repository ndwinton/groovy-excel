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
	
	static Font combine(Font self, Font other) {
		Font combined = combineWithoutSettingWorkbook(self, other)
		combined.workbook = self.workbook
		combined
	}
	
	private static Font combineWithoutSettingWorkbook(Font self, Font other) {
		Workbook workbook = self.workbook
		Font base = workbook.getFontAt(0 as short)
		def combined = [:]
		['boldweight', 'color', 'fontHeight',
			'fontName', 'italic', 'strikeout', 'typeOffset', 'underline'].each { attr ->
			combined[attr] = self[attr]
			if (other[attr] != base[attr]) {
				combined[attr] = other[attr]
			}
		}
		// Magic ahead! Spread the values of the (ordered) Map over the findFont arguments
		Font combinedFont = workbook.findFont(*(combined.collect { k, v -> v } ))
		if (!combinedFont) {
			combinedFont = workbook.createFont()
			combined.each { attribute, value ->
				combinedFont[attribute] = value
			}
		}
		combinedFont
	}
	
	static Font enhance(Font font, Workbook workbook) {
		font.metaClass {
			setWorkbook = { Workbook wb ->
				delegate.metaClass.getWorkbook = { -> wb }
			}
			
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
			
			combine = { Font other ->
				enhance(FontEnhancer.combineWithoutSettingWorkbook(delegate, other), workbook)
			}
		}
		font
	}
}
