package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet

import org.junit.Before;
import org.junit.Test;
import org.junit.Ignore

class ExcelBuilderTest {
	static final String TEST_FILE = "test.xlsx"
	ExcelBuilder builder
	
	@Before
	public void setUp() {
		builder = new ExcelBuilder()
	}
	
	@Test
	public void shouldBeABleToConstructWithNoArguments() {
		builder = new ExcelBuilder()
		assert builder instanceof BuilderSupport
	}
	
	@Test
	public void builderShouldContainAReadonlyWorkbookProperty() {
		assert builder.workbook
		try {
			builder.workbook = null
			fail("Shouldn't be able to set read-only property")
		}
		catch (ReadOnlyPropertyException e) {
			// OK
		}
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldThrowExceptionForUnknownBuilderContentType() {
		builder.doesNotExist()
	}
	
	@Test
	public void shouldBeABleToCreateUnnamedWorksheets() {
		Sheet s1 = builder.sheet()
		assert s1.sheetName == 'Sheet1'
		Sheet s2 = builder.sheet()
		assert s2.sheetName == 'Sheet2'
	}
	
	@Test
	public void shouldBeAbleToCreateNamedWorksheets() {
		Sheet s1 = builder.sheet("First")
		assert s1.sheetName == 'First'
		Sheet s2 = builder.sheet("Second")
		assert s2.sheetName == 'Second'
	}
	
	@Test
	public void shouldBeAbleToAccessCreatedSheetsByName() {
		builder {
			sheet()
			sheet('Second')
			sheet()
		}
		assert builder.sheets['Sheet1'].sheetName == 'Sheet1'
		assert builder.sheets['Second'].sheetName == 'Second'
		assert builder.sheets['Sheet2'].sheetName == 'Sheet2'
	}
	
	@Test
	public void shouldBeAbleToCreateSheetWithNameAttribute() {
		Sheet s = builder.sheet(name: 'My Name')
		assert s.sheetName == 'My Name'
	}
	
	@Test
	public void shouldBeAbleToSetASheetAsActive() {
		builder {
			sheet('one')
			sheet('two', active: true)
			sheet('three')
		}
		assert !builder.sheets.one.active
		assert builder.sheets.two.active
		assert !builder.sheets.three.active
		
		builder.sheets.three.active = true
		assert !builder.sheets.one.active
		assert !builder.sheets.two.active
		assert builder.sheets.three.active
	}

	
	@Test
	public void shouldBeAbleToSetASheetAsHidden() {
		builder {
			sheet('one')
			sheet('two', hidden: true)
			sheet('three')
		}
		assert !builder.sheets.one.hidden
		assert builder.sheets.two.hidden
		assert !builder.sheets.three.hidden
		
		builder.sheets.three.hidden = true
		assert builder.sheets.three.hidden
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldOnlyBeAbleToCreateASheetDirectlyFromABuilder() {
		builder {
			sheet() {
				sheet()
			}
		}
	}
	
	@Test
	public void shouldBeAbleToCreateARowWithAPreviouslyDefinedSheet() {
		Row r1, r2, r3
		builder {
			sheet('one') {
				r1 = row()
			}
		}
		builder.sheet('two') {
			r2 = row()
		}
		r3 = builder.row()
		
		assert r1 && builder.sheets.one.any { it == r1 }
		assert r2 && builder.sheets.two.any { it == r2 }
		assert r3 && builder.sheets.two.any { it == r3 }
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldNotBeAbleToCreateARowWithoutAPreviouslyDefinedSheet() {
		Row r
		builder {
			row()
		}
	}
	
	@Test
	public void shouldBeAbleToCreateCellsWithinARow() {
		Row r
		Date now = new Date()
		builder {
			sheet {
				r = row {
					cell(1)
					cell(2.3)
					cell('four')
					cell("${5}")
					cell(null)
					cell(true)
					cell(now)
				}
			}
		}
		
		Cell[] cells = r.collect {it}
		assert cells.size() == 7
		assert cells[0].getNumericCellValue() == 1
		assert cells[1].getNumericCellValue() == 2.3
		assert cells[2].getStringCellValue() == 'four'
		assert cells[3].getStringCellValue() == '5'
		assert cells[4].getStringCellValue() == ''
		assert cells[5].getBooleanCellValue() == true
		assert cells[6].getDateCellValue() == now
	}
	
	@Test
	public void shouldBeAbleToCreateCellsAsSpecificPositionsWithinARow() {
		Row r
		Date now = new Date()
		builder {
			sheet {
				r = row {
					cell(1, column: 10)
					cell(2.3, column: 2)
					cell('four', column: 4)
					cell("${5}", column: 6)
					cell(null, column: 20)
					cell(true, column: 15)
					cell(now, column: 1)
				}
			}
		}
		
		assert r.firstCellNum == 1
		assert r.lastCellNum == 21
				
		Cell[] cells = r.collect {it}
		assert cells.size() == 7
		
		assert cells[4].getNumericCellValue() == 1
		assert cells[4].columnIndex == 10
		
		assert cells[1].getNumericCellValue() == 2.3
		assert cells[1].columnIndex == 2
		
		assert cells[2].getStringCellValue() == 'four'
		assert cells[2].columnIndex == 4
		
		assert cells[3].getStringCellValue() == '5'
		assert cells[3].columnIndex == 6
		
		assert cells[6].getStringCellValue() == ''
		assert cells[6].columnIndex == 20

		assert cells[5].getBooleanCellValue() == true
		assert cells[5].columnIndex == 15
		
		assert cells[0].getDateCellValue() == now
		assert cells[0].columnIndex == 1
	}
	
	
	@Test
	public void newCellsWithImplicitColumnShouldFollowLastExplicitCell() {
		Row r
		Date now = new Date()
		builder {
			sheet {
				r = row {
					cell(1, column: 10)
					cell(2.3, column: 2)
					cell(4)
				}
			}
		}
		
		assert r.firstCellNum == 2
		assert r.lastCellNum == 11
				
		Cell[] cells = r.collect {it}
		assert cells.size() == 3
		
		assert cells[1].getNumericCellValue() == 4
		assert cells[1].columnIndex == 3
	}
	
	@Test
	public void shouldBeAbleToCreateCellsFromRowArgument() {
		Row r
		Date now = new Date()
		builder {
			sheet {
				r = row([1, 2.3, 'four', "${5}", null, true, now])
			}
		}
		
		Cell[] cells = r.collect {it}
		assert cells.size() == 7
		assert cells[0].getNumericCellValue() == 1
		assert cells[1].getNumericCellValue() == 2.3
		assert cells[2].getStringCellValue() == 'four'
		assert cells[3].getStringCellValue() == '5'
		assert cells[4].getStringCellValue() == ''
		assert cells[5].getBooleanCellValue() == true
		assert cells[6].getDateCellValue() == now
	}
	
	@Test
	public void cellPositionShouldResetAtTheStartOfANewRow() {
		Row r1, r2
		Date now = new Date()
		builder {
			sheet {
				r1 = row([1, 2.3, 'four', "${5}", null, true, now])
				r2 = row([1, 2.3, 'four', "${5}", null, true, now])
			}
		}
		assert r1.firstCellNum == 0
		assert r1.lastCellNum == 7
		assert r2.firstCellNum == 0
		assert r2.lastCellNum == 7
	}

	@Test
	public void laterCellValuesShouldOverridePreviousOnes() {
		Row r
		Date now = new Date()
		builder {
			sheet {
				r = row([1, 2.3, 'four', "${5}", null, true, now]) {
					cell(column: 2, 4)
					cell(column: 4, 'Hello')
				}
			}
		}
		
		Cell[] cells = r.collect {it}
		assert cells.size() == 7
		assert cells[0].getNumericCellValue() == 1
		assert cells[1].getNumericCellValue() == 2.3
		assert cells[2].getNumericCellValue() == 4
		assert cells[3].getStringCellValue() == '5'
		assert cells[4].getStringCellValue() == 'Hello'
		assert cells[5].getBooleanCellValue() == true
		assert cells[6].getDateCellValue() == now
	}
	
	@Test
	public void shouldBeAbleToCreateACellWithExplicitPositionAndValue() {
		Cell c
		builder {
			sheet {
				c = cell(row: 2, column: 10, value: 42)
			}
		}
		assert c.rowIndex == 2
		assert c.columnIndex == 10
		assert c.getNumericCellValue() == 42
	}
	
	@Test
	public void shouldBeAbleToCreateACellWithExplicitPositionInAnExistingRow() {
		Row r
		builder {
			sheet {
				r = row([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]) {
					cell(row: 0, column: 5, value: 42)
				}
			}
		}
		assert r.collect { it.getNumericCellValue() } == [1, 2, 3, 4, 5, 42, 7, 8, 9, 10]
	}
	
	@Test
	public void shouldBeAbleToCreateACellWithImplictPositionAndNoPreExistingRow() {
		Cell c
		builder {
			sheet {
				c = cell(42)
			}
		}
		assert c.rowIndex == 0
		assert c.columnIndex == 0
		assert c.getNumericCellValue() == 42
	}
	
	@Test
	public void shouldBeAbleToCreateRowAtAnExplicitPosition() {
		Row r1, r2, r3, r4
		builder {
			sheet {
				r1 = row([1, 2, 3])
				r2 = row([4, 5, 6], row: 3)
				r3 = row([7, 8, 9], row: 1)
				r4 = row(['a', 'b', 'c'])
			}
		}
		assert r1.rowNum == 0
		assert r2.rowNum == 3
		assert r3.rowNum == 1
		assert r4.rowNum == 2
	}
	
	@Test
	public void shouldBeAbleToDefineANamedFont() {
		Font f
		builder {
			f = font('italic') {
				italic = true 
			}
		}
		assert f.italic
		assert builder.fonts['italic'].italic
	}
	
	@Test
	public void shouldBeAbleToDefineNamedFontWithAllStandardAttributes() {
		builder {
			font('everything') {
				boldweight = Font.BOLDWEIGHT_BOLD
				charSet = Font.ANSI_CHARSET
				color = Font.COLOR_RED
				fontHeightInPoints = 12 as short
				fontName = 'Arial'
				italic = true
				strikeout = true
				typeOffset = Font.SS_SUPER
				underline = Font.U_SINGLE
			}
		}
		Font f = builder.fonts['everything']
		assert f.boldweight == Font.BOLDWEIGHT_BOLD
		assert f.color == Font.COLOR_RED
		assert f.fontHeight == (12 * 20)
		assert f.fontName == 'Arial'
		assert f.italic
		assert f.strikeout
		assert f.typeOffset == Font.SS_SUPER
		assert f.underline == Font.U_SINGLE
	}
	
	@Test
	public void shouldBeAbleToDefineNamedFontWithMetaClassAttributes() {
		builder {
			font('everything') {
				bold = true
				fontHeightInPoints = 12.5
			}
		}
		Font f = builder.fonts['everything']
		assert f.bold
		assert f.boldweight == Font.BOLDWEIGHT_BOLD
		assert f.fontHeight == (12.5 * 20)
		assert f.fontHeightInPoints == 12.5
	}
	
	@Test
	public void shouldBeAbleToDefineNamedStyle() {
		def s
		builder {
			s = style('style1')
		}
		assert s instanceof CellStyle
		assert builder.styles['style1']
	}
	
	@Test
	public void shouldBeAbleToDefineNamedStyleWithAllStandardAttributes() {
		CellStyle s
		builder {
			font('bold') {
				bold = true
			}
			
			s = style('omnistyle') {
				alignment = CellStyle.ALIGN_CENTER
				borderBottom = CellStyle.BORDER_DASH_DOT
				borderLeft = CellStyle.BORDER_DASHED
				borderRight = CellStyle.BORDER_DOTTED
				borderTop = CellStyle.BORDER_MEDIUM
				bottomBorderColor = IndexedColors.AQUA.index
				// dataFormat - handled in overloaded section
				fillForegroundColor = IndexedColors.BLACK.index	// Set FG before BG according to docs
				fillBackgroundColor = IndexedColors.BLUE.index
				fillPattern = CellStyle.BIG_SPOTS
				font = builder.fonts['bold']
				hidden = true
				indention = 2
				leftBorderColor = IndexedColors.BLUE_GREY.index
				locked = true
				rightBorderColor = IndexedColors.BRIGHT_GREEN.index
				rotation = 90
				// shrinkToFit = true // Not supported?
				topBorderColor = IndexedColors.BROWN.index
				verticalAlignment = CellStyle.VERTICAL_CENTER
				wrapText = true
			}
		}

		s.with {
			assert alignment == CellStyle.ALIGN_CENTER
			assert borderBottom == CellStyle.BORDER_DASH_DOT
			assert borderLeft == CellStyle.BORDER_DASHED
			assert borderRight == CellStyle.BORDER_DOTTED
			assert borderTop == CellStyle.BORDER_MEDIUM
			assert bottomBorderColor == IndexedColors.AQUA.index
			// dataFormat - handled in overloaded section
			assert fillForegroundColor == IndexedColors.BLACK.index	// Set FG before BG according to docs
			assert fillBackgroundColor == IndexedColors.BLUE.index
			assert fillPattern == CellStyle.BIG_SPOTS
			assert font == builder.fonts['bold']
			assert fontIndex == builder.fonts['bold'].index
			assert hidden
			assert indention == 2
			assert leftBorderColor == IndexedColors.BLUE_GREY.index
			assert locked
			assert rightBorderColor == IndexedColors.BRIGHT_GREEN.index
			assert rotation == 90
			//assert shrinkToFit
			assert topBorderColor == IndexedColors.BROWN.index
			assert verticalAlignment == CellStyle.VERTICAL_CENTER
			assert wrapText
		}
	}
	
	@Test
	public void shouldBeAbleToDefineNamedStyleWithMetaClassAttributes() {
		CellStyle s
		builder {
			font('bold') {
				bold = true
			}
			
			s = style('omnistyle') {
				dataFormatString = "0.00%"
			}
		}

		assert s.dataFormatString == "0.00%"
		assert s.getDataFormatString() == "0.00%"
	}

	@Test
	public void fontDefinitionsShouldCreateMatchingStyles() {
		Font f
		builder {
			f = font('bold') {
				bold = true
			}
		}
		assert builder.styles['bold'].font == f
	}
	
	@Test
	public void fontDefinitionShouldNotOverrideExistingStyle() {
		Font f
		builder {
			style('foo')
			f = font('foo') {
				italic = true
			}
		}
		assert builder.styles['foo'].font != f 
	}
	
	@Test
	public void shouldBeAbleToApplyStyleToACell() {
		Cell c
		builder {
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			sheet {
				c = cell('I am centred', style: 'centred')
			}
		}
		assert c.cellStyle.alignment == CellStyle.ALIGN_CENTER
	}
	
	@Test
	public void shouldBeAbleToApplyMultipleStylesToACell() {
		Cell c1, c2
		builder {
			font('bold') {
				bold = true
			}
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			sheet {
				c1 = cell('I am centred and bold', style: ['centred', 'bold'])
				c2 = cell('So am I', style: ['centred', 'bold'])
			}
		}
		assert c1.cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert c1.cellStyle.font == builder.fonts['bold']
		assert c1.cellStyle == builder.styles['centred+bold']
		
		assert c2.cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert c2.cellStyle.font == builder.fonts['bold']
		assert c2.cellStyle == builder.styles['centred+bold']
	}
	
	@Test
	public void shouldBeAbleToApplyStyleToARow() {
		Row r
		builder {
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			sheet {
				r = row([1, 2, 'Hello'], style: 'centred')
			}
		}
		assert r[0].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[1].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[2].cellStyle.alignment == CellStyle.ALIGN_CENTER
	}
	
	@Test
	public void shouldBeAbleToOverrideStyleForCellWithinARow() {
		Row r
		builder {
			font('bold') {
				bold = true
			}
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			sheet {
				r = row([1, 2, 'Hello'], style: 'centred') {
					cell(column: 1, style: 'bold')
				}
			}
		}
		assert r[0].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[0].numericCellValue == 1
		assert r[1].cellStyle.alignment == CellStyle.ALIGN_GENERAL
		assert r[1].cellStyle.font == builder.fonts['bold']
		assert r[1].numericCellValue == 2
		assert r[2].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[2].stringCellValue == 'Hello'
	}
	
	@Test
	public void rowStyleShouldApplyToNewCellsCreatedWithinTheRow() {
		Row r
		builder {
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			sheet {
				r = row([1, 2, 'Hello'], style: 'centred') {
					cell('Me too', column: 3)
				}
			}
		}
		assert r.rowStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[0].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[1].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[2].cellStyle.alignment == CellStyle.ALIGN_CENTER
		assert r[3].cellStyle.alignment == CellStyle.ALIGN_CENTER
	}

	@Test(expected=IllegalArgumentException)
	public void shouldThrowExceptionForUnknownStyle() {
		Cell c
		builder {
			sheet {
				c = cell('Bad style', style: 'undefined')
			}
		}
	}
	
	@Test
	public void shouldBeAbleToSpecifyHeightOfARowInPoints() {
		Row r
		builder {
			sheet {
				r = row([1, 2, 'Hello'], height: 20.5)
			}
		}
		assert r.heightInPoints == 20.5
	}
	
	@Test
	public void shouldBeAbleToSpecifyHeightOfACellInPoints() {
		Cell c
		builder {
			sheet {
				c = cell(row: 1, column:0, height: 20.5)
			}
		}
		assert c.row.heightInPoints == 20.5
	}
	
	@Test
	public void settingCellHeightShouldHaveNoEffectIfRowIsAlreadyLarger() {
		Row r
		builder {
			sheet {
				r = row([1, 2, 'Hello'], height: 50) {
					cell(height: 30)
				}
			}
		}
		assert r.heightInPoints == 50
	}
	
	@Test
	public void shouldBeAbleToForceLesserHeightByUsingNegativeOrZeroValue() {
		Row r1, r2
		builder {
			sheet {
				r1 = row([1, 2, 'Hello'], height: 50) {
					cell(height: -30)
				}
				r2 = row(height: 0)
			}
		}
		assert r1.heightInPoints == 30
		assert r2.heightInPoints == 0
	}
	
	@Test
	public void shouldBeAbleToSetCellWidthInChars() {
		Sheet s
		builder {
			s = sheet {
				cell(row: 1, column:1, width: 50)
			}
		}
		assert s.getColumnWidth(1) == 50 * 256
	}
	
	@Test
	public void settingCellWidthShouldHaveNoEffectIfCellIsAlreadyLarger() {
		Sheet s
		builder {
			s = sheet {
				cell(row: 0, column: 1, width: 100)
				cell(row: 1, column: 1, width: 50)
			}
		}
		assert s.getColumnWidth(1) == 100 * 256
	}
	
	@Test
	public void shouldBeAbleToForceLesserWidthByUsingNegativeOrZeroValue() {
		Sheet s
		builder {
			s = sheet {
				cell(row: 0, column: 1, width: 100)
				cell(row: 1, column: 1, width: -50)
				cell(column: 0, width: 0)
			}
		}
		assert s.getColumnWidth(0) == 0
		assert s.getColumnWidth(1) == 50 * 256
	}
	
	@Test
	public void shouldBeAbleToSetCellWidthInCharsForARow() {
		Sheet s
		builder {
			s = sheet {
				row([1, 2, 'Hello'], width: 50)
			}
		}
		(0 .. 2).each {
			assert s.getColumnWidth(it) == 50 * 256
		}
	}
	
	@Test
	public void rowCellWidthShouldPropagateToNewCells() {
		Sheet s
		builder {
			s = sheet {
				row([1, 2, 'Hello'], width: 50) {
					cell(42)
				}
			}
		}
		(0 .. 3).each {
			assert s.getColumnWidth(it) == 50 * 256
		}
	}
	
	@Test
	public void shouldBeAbleToCombineEverything() {
		builder {
			font('bold') {
				bold = true
			}
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			style('uk-date') {
				dataFormatString = 'dd/mm/yyyy'
			}
			sheet {
				row([1, 2, 'Hello'], style: 'centred') {
					cell(column: 1, style: ['bold', 'centred'])
					cell(column: 4, 'Me too')
				}
				row([5, 6, 7], style: 'bold')
				row(['Today', new Date()]) {
					cell(column: 1, style: 'uk-date')
				}
			}
		}
		builder.workbook.write(new FileOutputStream(new File(TEST_FILE)))
	}
}