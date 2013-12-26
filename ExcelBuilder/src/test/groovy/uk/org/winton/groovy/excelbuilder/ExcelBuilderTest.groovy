package uk.org.winton.groovy.excelbuilder;

import static org.junit.Assert.*;

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
		
		assert r1 && builder.sheets.one.rows.any { it == r1 }
		assert r2 && builder.sheets.two.rows.any { it == r2 }
		assert r3 && builder.sheets.two.rows.any { it == r3 }
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void shouldNotBeAbleToCreateARowWithoutAPreviouslyDefinedSheet() {
		Row r
		builder {
			row()
		}
	}
}
