package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class SheetEnhancer {

	static void setActive(Sheet self, boolean active) {
		Workbook wb = self.workbook
		if (active) {
			wb.activeSheet = wb.getSheetIndex(self)
		}
	}
	
	static boolean getActive(Sheet self) {
		self.workbook.activeSheetIndex == self.workbook.getSheetIndex(self)
	}
	
	static void setHidden(Sheet self, boolean hidden) {
		Workbook wb = self.workbook
		wb.setSheetHidden(wb.getSheetIndex(self), hidden)
	}
	
	static boolean getHidden(Sheet self) {
		self.workbook.isSheetHidden(self.workbook.getSheetIndex(self))
	}
	
	static void setVeryHidden(Sheet self, boolean hidden) {
		Workbook wb = self.workbook
		wb.setSheetHidden(wb.getSheetIndex(self), hidden ? Workbook.SHEET_STATE_VERY_HIDDEN : Workbook.SHEET_STATE_VISIBLE)
	}
	
	static boolean getVeryHidden(Sheet self) {
		self.workbook.isSheetVeryHidden(self.workbook.getSheetIndex(self))
	}
	
	static void setColumnWidthInChars(Sheet self, int index, Number width) {
		self.setColumnWidth(index, Math.min((width * 256) as int, 255 * 256))
	}
	
	static Number getColumnWidthInChars(Sheet self, int index) {
		self.getColumnWidth(index) / 256.0
	}

	static void setDefaultColumnWidthInChars(Sheet self, Number width) {
		self.setDefaultColumnWidth(Math.min(Math.ceil(width) as int, 255))
	}
	
	static Number getDefaultColumnWidthInChars(Sheet self) {
		self.getDefaultColumnWidth()
	}
	
	static Sheet enhance(Sheet sheet) {
		sheet.metaClass {
			setActive = { boolean active ->
				SheetEnhancer.setActive(delegate, active)
			}
			
			getActive = { ->
				SheetEnhancer.getActive(delegate)
			}
			
			setHidden = { boolean hidden ->
				SheetEnhancer.setHidden(delegate, hidden)
			}
			
			getHidden = { ->
				SheetEnhancer.getHidden(delegate)
			}

			setVeryHidden = { boolean hidden ->
				SheetEnhancer.setVeryHidden(delegate, hidden)
			}
			
			getVeryHidden = { ->
				SheetEnhancer.getVeryHidden(delegate)
			}
			
			setColumnWidthInChars = { int column, Number width ->
				SheetEnhancer.setColumnWidthInChars(delegate, column, width)
			}
			
			getColumnWidthInChars = { int column ->
				SheetEnhancer.getColumnWidthInChars(delegate, column)
			}
			
			setDefaultColumnWidthInChars = { Number width ->
				SheetEnhancer.setDefaultColumnWidthInChars(delegate, width)
			}
			
			getDefaultColumnWidthInChars = {
				SheetEnhancer.getDefaultColumnWidthInChars(delegate)
			}
		}
		sheet
	}
	
}
