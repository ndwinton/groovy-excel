package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*;

import org.apache.poi.ss.usermodel.Cell
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
}
