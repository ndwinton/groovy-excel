package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Workbook

class CellStyleEnhancer {
	private static def workbookMap = Collections.synchronizedMap([:])
	
	static void setWorkbook(CellStyle self, Workbook workbook) {
		workbookMap[self] = workbook
	}
	
	static Workbook getWorkbook(CellStyle self) {
		workbookMap[self]
	}
	
	static void setDataFormatString(CellStyle self, String formatString) {
		DataFormat fmt = self.workbook.creationHelper.createDataFormat()
		self.setDataFormat(fmt.getFormat(formatString))
	}
	
	static CellStyle enhance(CellStyle style, Workbook workbook) {
		setWorkbook(style, workbook)
		style.metaClass {
			getWorkbook = { ->
				workbookMap[style]
			}
			
			setDataFormatString = { String str ->
				CellStyleEnhancer.setDataFormatString(delegate, str)
			}
			
			getDataFormatString = { ->
				delegate.getDataFormatString()
			}
		}
		style
	}
}
