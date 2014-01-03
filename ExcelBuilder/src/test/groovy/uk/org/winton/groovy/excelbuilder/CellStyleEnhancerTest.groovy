package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Before;
import org.junit.Test;

class CellStyleEnhancerTest {
	Workbook wb
	CellStyle s1
	
	@Before
	public void setUp() throws Exception {
		wb = new XSSFWorkbook()
		s1 = wb.createCellStyle()
		
	}

	@Test
	public void shouldBeAbleToSetWorkbookUsingClassAsCategory() {
		use (CellStyleEnhancer) {
			s1.workbook = wb
			assert s1.workbook == wb
		}
	}
	
	@Test
	public void shouldBeAbleToSetDataFormatStringUsingClassAsCategory() {
		use (CellStyleEnhancer) {
			s1.workbook = wb
			s1.dataFormatString = "0.00%"
			assert s1.dataFormatString == "0.00%"
		}
	}
	
	@Test
	public void shouldBeAbleToCombinedStylesToGenerateANewOne() {
		use (CellStyleEnhancer) {
			s1.workbook = wb
			s1.alignment = CellStyle.ALIGN_CENTER
			s1.font = wb.createFont()
			s1.font.italic = true
			
			CellStyle s2 = wb.createCellStyle()
			s2.workbook = wb
			s2.borderBottom = CellStyle.BORDER_DASH_DOT_DOT
			s2.font = wb.createFont()
			s2.font.fontName = 'Courier'
			
			CellStyle s3 = s1.combine(s2)
			assert s3.alignment == CellStyle.ALIGN_CENTER
			assert s3.borderBottom == CellStyle.BORDER_DASH_DOT_DOT
			assert s3.font.italic
			assert s3.font.fontName == 'Courier' 
		}
	}
	
	@Test
	public void shouldBeAbleToCombineMultipleStyles() {
		use (CellStyleEnhancer) {
			s1.workbook = wb
			s1.alignment = CellStyle.ALIGN_CENTER
			s1.font = wb.createFont()
			s1.font.italic = true
			
			CellStyle s2 = wb.createCellStyle()
			s2.workbook = wb
			s2.borderBottom = CellStyle.BORDER_DASH_DOT_DOT
			s2.font = wb.createFont()
			s2.font.fontName = 'Courier'
			
			CellStyle s3 = wb.createCellStyle()
			s3.verticalAlignment = CellStyle.VERTICAL_TOP
			
			CellStyle s4 = s1.combine(s2, s3)
			assert s4.alignment == CellStyle.ALIGN_CENTER
			assert s4.borderBottom == CellStyle.BORDER_DASH_DOT_DOT
			assert s4.font.italic
			assert s4.font.fontName == 'Courier'
			assert s4.verticalAlignment == CellStyle.VERTICAL_TOP
		}
	}

	@Test
	public void shouldBeAbleToEnhanceInstanceMetaClass() {
		CellStyleEnhancer.enhance(s1, wb)
		assert s1.workbook == wb
		s1.dataFormatString = "0.00%"
		assert s1.dataFormatString == "0.00%"
		assert s1.getDataFormatString() == "0.00%"
	}
	
	@Test
	public void shouldBeAbleToCombinedEnhancedCellStyleWithNonEnhancedToGiveNewEnhancedCellStyle() {
		CellStyleEnhancer.enhance(s1, wb)
		s1.dataFormatString = "0.00%"

		CellStyle s2 = wb.createCellStyle()
		s2.alignment = CellStyle.ALIGN_CENTER
		
		CellStyle s3 = s1.combine(s2)
		assert s3.workbook == wb
		assert s3.alignment == CellStyle.ALIGN_CENTER
		assert s3.dataFormatString == "0.00%"
	}
}
