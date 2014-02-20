package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * <p>SheetEnhancer enriches the {@link Sheet} class. It can be used as either a
 * category class (with <code>use (SheetEnhancer) { ... }</code>)
 * or a Sheet instance can have permanent meta-class modifications made with
 * the {@link #enhance(Sheet)} method.</p>
 *
 * <p>The class adds the following properties:
 * <dl>
 * <dt><b>active</b></dt>
 * 	<dd>A boolean value that marks the sheet as active.</dd>
 * <dt><b>hidden</b></dt>
 * 	<dd>A boolean value that marks the sheet as hidden.</ddt>
 * <dt><b>veryHidden</b></dt>
 * 	<dd>A boolean value that marks the sheet as very hidden
 * (see {@link org.apache.poi.ss.usermodel.Workbook#setSheetHidden(int, int)} for an
 * explanation of this).</dd>
 * </dl>
 * It also adds methods to get and set column widths in characters (including fractional values)
 * rather than the normal 256ths of a character. Finally {@link #enhance(Sheet)} adds permanent meta-class
 * modifications to an instance.
 * </p>
 *
 * @author Neil Winton
 *
 */
class SheetEnhancer {

	/**
	 * Marks the sheet as active (or not) within its workbook.
	 * 
	 * @param self the sheet to modify
	 * @param active true to make the sheet active, false to make it inactive
	 */
	static void setActive(Sheet self, boolean active) {
		Workbook wb = self.workbook
		if (active) {
			wb.activeSheet = wb.getSheetIndex(self)
		}
	}
	
	/**
	 * Returns the current active state of the worksheet.
	 * 
	 * @param self the sheet for which to retrieve the state
	 * @return true if the sheet is marked as active, false otherwise
	 */
	static boolean getActive(Sheet self) {
		self.workbook.activeSheetIndex == self.workbook.getSheetIndex(self)
	}
	
	/**
	 * Marks the sheet as hidden (or not) within its workbook.
	 * 
	 * @param self the sheet to modify
	 * @param hidden true to make the sheet hidden, false to make it visible
	 */
	static void setHidden(Sheet self, boolean hidden) {
		Workbook wb = self.workbook
		wb.setSheetHidden(wb.getSheetIndex(self), hidden)
	}
	
	/**
	 * Returns the current visibility state of the worksheet.
	 * 
	 * @param self the sheet for which to retrieve the state
	 * @return true if the sheet is hidden, false otherwise
	 */
	static boolean getHidden(Sheet self) {
		self.workbook.isSheetHidden(self.workbook.getSheetIndex(self))
	}
	
	/**
	 * Marks the sheet as "very hidden" (or not) within its workbook.
	 * @see org.apache.poi.ss.usermodel.Workbook#setSheetHidden(int, int)
	 * 
	 * @param self the sheet to modify
	 * @param hidden true to make the sheet "very hidden", false to make it visible
	 */
	static void setVeryHidden(Sheet self, boolean hidden) {
		Workbook wb = self.workbook
		wb.setSheetHidden(wb.getSheetIndex(self), hidden ? Workbook.SHEET_STATE_VERY_HIDDEN : Workbook.SHEET_STATE_VISIBLE)
	}
	
	/**
	 * Returns the current "very hidden" state of the worksheet.
	 * 
	 * @param self the sheet for which to retrieve the state
	 * @return true if the sheet is very hidden, false otherwise
	 */
	static boolean getVeryHidden(Sheet self) {
		self.workbook.isSheetVeryHidden(self.workbook.getSheetIndex(self))
	}
	
	/**
	 * Sets the width of a column in characters (including fractional values).
	 * There is a maximum width of 255 characters. Values larger than this
	 * are truncated to this value.
	 * 
	 * @param self the sheet containing the column
	 * @param index the index of the column
	 * @param width the width of the column
	 */
	static void setColumnWidthInChars(Sheet self, int index, Number width) {
		self.setColumnWidth(index, Math.min((width * 256) as int, 255 * 256))
	}
	
	/**
	 * Retrieves the width of a column in characters.
	 * 
	 * @param self the sheet containing the column
	 * @param index the index of the column
	 * @return the width of the specified column in characters
	 */
	static Number getColumnWidthInChars(Sheet self, int index) {
		self.getColumnWidth(index) / 256.0
	}

	/**
	 * Sets the default width of a columns in a sheet in characters
	 * (including fractional values). There is a maximum width of 255
	 * characters. Values larger than this are truncated to this value.
	 * 
	 * @param self the Sheet instance
	 * @param width the default width of columns in characters
	 */
	static void setDefaultColumnWidthInChars(Sheet self, Number width) {
		self.setDefaultColumnWidth(Math.min(Math.ceil(width) as int, 255))
	}
	
	/**
	 * Retrieves the default column width in characters.
	 * 
	 * @param self the Sheet instance
	 * @return the default width of columns in characters
	 */
	static Number getDefaultColumnWidthInChars(Sheet self) {
		self.getDefaultColumnWidth()
	}
	
	/**
	 * This method makes permanent modifications to a Sheet instance's meta-class
	 * to make all of the other methods (and properties) in the SheetEnhancer
	 * class directly available to the instance.
	 *
	 * @param sheet the Sheet instance to enhance
	 * @return the enhanced CellStyle instance
	 */
	
	static Sheet enhance(Sheet sheet) {
		sheet.metaClass {
			setActive = { boolean active ->
				SheetEnhancer.setActive(delegate, active)
			}
			
			getActive = { ->
				SheetEnhancer.getActive(delegate)
			}
			
			setHidden = { boolean hidden ->
				SheetEnhancer.setHidden(delegate, hidden)
			}
			
			getHidden = { ->
				SheetEnhancer.getHidden(delegate)
			}

			setVeryHidden = { boolean hidden ->
				SheetEnhancer.setVeryHidden(delegate, hidden)
			}
			
			getVeryHidden = { ->
				SheetEnhancer.getVeryHidden(delegate)
			}
			
			setColumnWidthInChars = { int column, Number width ->
				SheetEnhancer.setColumnWidthInChars(delegate, column, width)
			}
			
			getColumnWidthInChars = { int column ->
				SheetEnhancer.getColumnWidthInChars(delegate, column)
			}
			
			setDefaultColumnWidthInChars = { Number width ->
				SheetEnhancer.setDefaultColumnWidthInChars(delegate, width)
			}
			
			getDefaultColumnWidthInChars = {
				SheetEnhancer.getDefaultColumnWidthInChars(delegate)
			}
		}
		sheet
	}
	
}
