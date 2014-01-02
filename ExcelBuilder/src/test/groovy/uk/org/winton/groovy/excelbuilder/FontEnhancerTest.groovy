package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*

import javax.swing.text.StyledEditorKit.BoldAction;

import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.Before
import org.junit.Test

class FontEnhancerTest {
	Workbook wb
	Font f1, f2
	
	@Before
	public void setUp() throws Exception {
		wb = new XSSFWorkbook()
		f1 = wb.createFont()
		f2 = wb.createFont()
	}

	@Test
	public void shouldBeAbleToSetWorkbookUsingClassAsCategory() {
		use (FontEnhancer) {
			f1.workbook = wb
			assert f1.workbook == wb
		}
	}
	
	@Test
	public void shouldBeAbleToSetBoldPropertyUsingClassAsCategory() {
		use (FontEnhancer) {
			assert !f1.bold
			assert f1.boldweight == Font.BOLDWEIGHT_NORMAL
			f1.bold = true
			assert f1.bold
			assert f1.boldweight == Font.BOLDWEIGHT_BOLD
		}
	}

	@Test
	public void shouldBeAbleToGetAndSetTheFontHeightInPointsUsingNonShortValues() {
		use (FontEnhancer) {
			f1.fontHeightInPoints = 20.5
			assert f1.fontHeightInPoints == 20.5
			assert f1.fontHeight == 410
			f1.fontHeightInPoints = 2000
			assert f1.fontHeightInPoints == Short.MAX_VALUE / 20.0
		}
	}
	
	@Test
	public void shouldBeAbleToCombineFontsToGenerateANewFont() {
		Font f3
		use (FontEnhancer) {
			f1.workbook = f2.workbook = wb
			f1.bold = true
			f1.fontName = 'Times'
			f1.fontHeightInPoints = 12
			f1.charSet = Font.SYMBOL_CHARSET
			f2.italic = true
			f2.fontName = 'Courier'
			f3 = f1.combine(f2)
			assert f3.bold
			assert f3.fontName == 'Courier'
			assert f3.fontHeightInPoints == 12
			assert f3.italic
			assert f3.charSet == Font.SYMBOL_CHARSET
			assert f3.workbook == f1.workbook
		}
	}

	@Test
	public void combiningFontsShouldReuseExistingFontsWhenAvailable() {
		Font f3 = wb.createFont()
		use (FontEnhancer) {
			f1.workbook = f2.workbook = f3.workbook = wb
			f1.bold = true
			f1.fontName = 'Times'
			f1.fontHeightInPoints = 12
			
			f2.italic = true
			f2.fontName = 'Courier'
			
			f3.bold = true
			f3.fontName = 'Courier'
			f3.fontHeightInPoints = 12
			f3.italic = true
			
			assert f1.combine(f2) == f3
		}
	}

	@Test
	public void shouldBeAbleToEnhanceInstanceMetaClass() {
		FontEnhancer.enhance(f1, wb)
		f1.bold = true
		f1.fontHeightInPoints = 20.5
		assert f1.workbook == wb
		assert f1.bold
		assert f1.fontHeightInPoints == 20.5
	}
	
	@Test
	public void shouldBeAbleToCombinedEnhancedFontWithNonEnhancedToGiveNewEnhancedFont() {
		FontEnhancer.enhance(f1, wb)
		f1.bold = true
		f1.fontHeightInPoints = 20.5
		
		f2.italic = true
		f2.fontName = 'Courier'
		
		Font f3 = f1.combine(f2)
		assert f3.fontName == 'Courier'
		assert f3.workbook == f1.workbook
		assert f3.fontHeightInPoints == 20.5
	}
}
