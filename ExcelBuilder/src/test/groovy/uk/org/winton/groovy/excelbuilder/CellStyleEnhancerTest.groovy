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
	public void shouldBeAbleToEnhanceInstanceMetaClass() {
		CellStyleEnhancer.enhance(s1, wb)
		assert s1.workbook == wb
		s1.dataFormatString = "0.00%"
		assert s1.dataFormatString == "0.00%"
		assert s1.getDataFormatString() == "0.00%"
		
	}
}
