package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet

import org.junit.Before;
import org.junit.Test;
import org.junit.Ignore

class ExcelBuilderTest {
	static final String TEST_FILE = "test.xlsx"
	static final String TEMPLATE_FILE_NAME = ExcelBuilderTest.classLoader.getResource("test-template.xlsx").file
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
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldThrowExceptionForUnknownBuilderContentTypeWithValue() {
		builder.doesNotExist('foo')
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldThrowExceptionForUnknownBuilderContentTypeWithAttributes() {
		builder.doesNotExist(foo: 42)
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldThrowExceptionForUnknownBuilderContentTypeWithAttributesAndValue() {
		builder.doesNotExist(123, foo: 42)
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
	public void shouldBeAbleToForceLesserCellHeightByUsingForceOption() {
		Row r1
		builder {
			sheet {
				r1 = row([1, 2, 'Hello'], height: 50) {
					cell(height: 30, force: true)
				}
			}
		}
		assert r1.heightInPoints == 30
	}
	
	@Test
	public void shouldBeAbleToSetCellWidthInChars() {
		Sheet s
		builder {
			s = sheet {
				cell(row: 1, column:1, width: 50)
			}
		}
		assert s.getColumnWidthInChars(1) == 50
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
		assert s.getColumnWidthInChars(1) == 100
	}
	
	@Test
	public void shouldBeAbleToForceLesserWidthByUsingForceOption() {
		Sheet s
		builder {
			s = sheet {
				cell(row: 0, column: 1, width: 100)
				cell(row: 1, column: 1, width: 50, force: true)
			}
		}
		assert s.getColumnWidthInChars(1) == 50
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
			assert s.getColumnWidthInChars(it) == 50
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
			assert s.getColumnWidthInChars(it) == 50
		}
	}
	
	@Test
	public void shouldBeAbleToSetDefaultWidthAndHeightForCellsInSheet() {
		Sheet s
		Row r
		builder {
			s = sheet(width: 50, height: 20) {
				r = row([1, 2, 'Hello'])
			}
		}
		assert s.defaultColumnWidthInChars == 50
		assert s.defaultRowHeightInPoints == 20
		(0 .. 3).each {
			assert s.getColumnWidthInChars(it) == 50
		}
		assert r.heightInPoints == 20
	}
	
	@Test
	public void shouldHaveASetOfPredefinedDateStyles() {
		assert builder.styles.collect {
			k, v -> k
		}.containsAll(['default-date', 'default-numeric', 'default-text', 'default-boolean',
			'iso-date', 'iso-datetime',
			'euro-date', 'euro-datetime',
			'us-date', 'us-datetime'])
	}
	
	@Test
	public void datesAndCalendarsShouldBeFormattedWithDefaultDateStyle() {
		Row r
		Calendar cal = new GregorianCalendar(2014, 0, 6, 12, 34, 56)
		builder {
			sheet {
				r = row([cal.time, cal])
			}
		}
		
		def fmt = new DataFormatter()
		assert fmt.formatCellValue(r[0]) == '2014/01/06 12:34:56'
		assert fmt.formatCellValue(r[1]) == '2014/01/06 12:34:56'
	}
	
	@Test
	public void shouldBePossibleToOverrideTheDefaultDateStyle() {
		Row r
		Calendar cal = new GregorianCalendar(2014, 0, 6, 12, 34, 56)
		builder.styles['default-date'] = builder.styles['euro-date']
		builder {
			sheet {
				r = row([cal.time, cal])
			}
		}
		
		def fmt = new DataFormatter()
		assert fmt.formatCellValue(r[0]) == '06/01/2014'
		assert fmt.formatCellValue(r[1]) == '06/01/2014'
	}
	
	@Test
	public void shouldBePossibleToOverrideTheDefaultNumericTextAndBooleanStyling() {
		Row r
		builder {
			font('bold') {
				bold = true
			}
			font('italic') {
				italic = true
			}
			font('strikeout') {
				strikeout = true
			}
			
			styles['default-text'] = styles['bold']
			styles['default-numeric'] = styles['italic']
			styles['default-boolean'] = styles['strikeout']

			sheet {
				r = row([42, 'Hello', true])
			}
		}
		assert r[0].cellStyle == builder.styles['italic']
		assert r[1].cellStyle == builder.styles['bold']
		assert r[2].cellStyle == builder.styles['strikeout']
	}
	
	@Test
	public void shouldBePossibleToSetFormulaeAndHaveThemEvaluated() {
		Row r
		builder {
			sheet {
				def amount = 3
				r = row(["NOT a formula",'=LEFT(A1,3)',"=RIGHT(A1,$amount)"])
			}
		}
		assert r[0].stringCellValue == 'NOT a formula'
		assert r[1].cellFormula == 'LEFT(A1,3)'
		assert r[1].stringCellValue == 'NOT'
		assert r[2].cellFormula == 'RIGHT(A1,3)'
		assert r[2].stringCellValue == 'ula'
	}
	
	@Test
	public void evaluatedFormulaCellShouldHaveCorrectValueTypeAndStyling() {
		Row r
		builder {
			sheet {
				font('bold') {
					bold = true
				}
				font('italic') {
					italic = true
				}
				font('strikeout') {
					strikeout = true
				}
				
				styles['default-text'] = styles['bold']
				styles['default-numeric'] = styles['italic']
				styles['default-boolean'] = styles['strikeout']
				r = row(["NOT a formula",'=LEFT(A1,3)','=LEN(A1)', '=A1="Foobar"'])
			}
		}
		assert r[0].stringCellValue == 'NOT a formula'
		assert r[1].stringCellValue == 'NOT'
		assert r[1].cellStyle == builder.styles['bold']
		assert r[2].numericCellValue == 13
		assert r[2].cellStyle == builder.styles['italic']
		assert !r[3].booleanCellValue
		assert r[3].cellStyle == builder.styles['strikeout']
	}
	
	@Test
	public void shouldBeAbleToForceValueToTextWithLeadingQuote() {
		Row r
		builder {
			sheet {
				r = row(["NOT a formula",'=LEFT(A1,3)',"'=LEN(A1)"])
			}
		}
		assert r[0].stringCellValue == 'NOT a formula'
		assert r[1].stringCellValue == 'NOT'
		assert r[2].stringCellValue == '=LEN(A1)'
	}
	
	@Test
	public void shouldBeAbleToCombineEverythingToGenerateNewFile() {
		builder {
			font('title') {
				bold = true
				color = IndexedColors.RED.index
			}
			style('centred') {
				alignment = CellStyle.ALIGN_CENTER
			}
			style('uk-date') {
				dataFormatString = 'dd/mm/yyyy'
			}
			sheet {
				row([1, 2, 'Hello'], style: ['centred', 'title'])
				row([5, 6, 7], style: 'centred')
				row(['Today', new Date()]) {
					cell(column: 1, style: 'uk-date', width: 20)
				}
			}
		}
		builder.workbook.write(new FileOutputStream(new File(TEST_FILE)))
	}
	
	@Test
	public void shouldBeAbleToConstructAnInstanceWithATemplateFile() {
		def templateFile = new File(TEMPLATE_FILE_NAME)
		builder = new ExcelBuilder(templateFile)
		assert builder
		assert builder.templateFile.canonicalPath == templateFile.canonicalPath
	}
	
	@Test
	public void shouldBeAbleToConstructAnInstanceWithATemplateFileName() {
		builder = new ExcelBuilder(TEMPLATE_FILE_NAME)
		assert builder
		assert builder.templateFile.canonicalPath == new File(TEMPLATE_FILE_NAME).canonicalPath
	}
	
	@Test
	public void workbookShouldContainTemplateFileContentsIfNoOtherModificationsAreMade() {
		builder = new ExcelBuilder(TEMPLATE_FILE_NAME)
		assert builder.workbook.numberOfSheets == 4
		Sheet data = builder.workbook.getSheet("Data")
		assert data
		data.getRow(0).getCell(0).getStringCellValue() == 'Alpha'
	}
	
	@Test
	public void usingNamedSheetWithTemplateShouldModifyThatSheet() {
		builder = new ExcelBuilder(TEMPLATE_FILE_NAME)
		Sheet data
		builder {
			data = sheet('Data') {
				cell('Hello', row: 0, column: 0)
				row([1, 2, 3], row: 1)
			}
		}
		assert data.getRow(0).getCell(0).stringCellValue == 'Hello'
		assert data.getRow(0).getCell(1).stringCellValue == 'Beta'
		assert data.getRow(1).getCell(2).numericCellValue == 3		
	}
	
	@Test
	public void sheetsInTemplateShouldAppearInBuilderSheetsProperty() {
		builder = new ExcelBuilder(TEMPLATE_FILE_NAME)
		assert builder.sheets['Data']?.sheetName == 'Data'
	}
	
	@Test
	public void newlyGeneratedSheetNamesShouldBeUnique() {
		builder = new ExcelBuilder(TEMPLATE_FILE_NAME)
		Sheet data
		builder {
			data = sheet {
				row([1, 2, 3])
			}
		}
		assert data.sheetName == 'Sheet3'
	}
	
}