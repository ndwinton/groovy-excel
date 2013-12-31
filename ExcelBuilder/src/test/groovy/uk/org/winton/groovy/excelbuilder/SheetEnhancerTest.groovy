package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.*

class SheetEnhancerTest {
	Workbook wb
	Sheet s1, s2
	
	@Before
	public void setUp() {
		wb = new XSSFWorkbook()
		s1 = wb.createSheet()
		s2 = wb.createSheet()
	}
	
	@Test
	public void shouldBeAbleToAddActivePropertyToSheetUsingClassAsCategory() {
		use (SheetEnhancer) {
			s1.active = true
			assert s1.active
			assert !s2.active
			assert wb.activeSheetIndex == wb.getSheetIndex(s1)
			s2.active = true
			assert !s1.active
			assert s2.active
			assert wb.activeSheetIndex == wb.getSheetIndex(s2)
		}
	}

	@Test
	public void shouldBeAbleToAddHiddenAndVeryHiddenPropertiesToSheetUsingClassAsCategory() {
		use (SheetEnhancer) {
			assert !s1.hidden
			assert !s1.veryHidden
			s1.hidden = true
			assert s1.hidden
			assert !s1.veryHidden
			assert wb.isSheetHidden(wb.getSheetIndex(s1))
			s1.veryHidden = true
			assert !s1.hidden
			assert s1.veryHidden
			assert wb.isSheetVeryHidden(wb.getSheetIndex(s1))
			s1.hidden = false
			assert !s1.hidden && !s1.veryHidden
		}
	}
	
	@Test
	public void shouldBeAbleToEnhanceInstanceMetaClass() {
		SheetEnhancer.enhance(s1)
		SheetEnhancer.enhance(s2)
		s1.active = true
		assert s1.active
		assert !s2.active
		assert wb.activeSheetIndex == wb.getSheetIndex(s1)
		s1.hidden = true
		s2.veryHidden = true
		assert s1.hidden
		assert wb.isSheetHidden(wb.getSheetIndex(s1))
		assert s2.veryHidden
		assert wb.isSheetVeryHidden(wb.getSheetIndex(s2))
	}
}
