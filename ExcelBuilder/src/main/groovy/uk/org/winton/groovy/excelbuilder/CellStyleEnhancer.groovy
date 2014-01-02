package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
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
		}
		style
	}
}
