package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.Font

class FontEnhancer {

	// Not actually necessary for XSSFFont class, which already has a getBold/setBold,
	// but it still works and also for HSSFFont
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
	
	static Font enhance(Font font) {
		font.metaClass {
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
		}
		font
	}
}
