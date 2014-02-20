package uk.org.winton.groovy.excelbuilder

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook

/**
 * <p>CellStyleEnhancer enriches the {@link org.apache.poi.ss.usermodel.CellStyle}
 * class. It can be used as either a category class (with
 * <code>use (CellStyleEnhancer) { ... }</code>)
 * or a CellStyle instance can have permanent meta-class modifications made with
 * the {@link #enhance(CellStyle, Workbook)} method.</p>
 * 
 * <p>The class adds the following properties:
 * <dl>
 * <dt>workbook</dt>
 * 	<dd>Holds the workbook within which this CellStyle was created. When using CellStyleEnhancer
 * 	as a category class, you must set the workbook before using the other methods.</dd>
 * <dt>dataFormatString</dt>
 * 	<dd>Adds a setter to match the existing getter, making the property writable.</ddt>
 * </dl>
 * It also adds the {@link #combine(CellStyle, CellStyle...)} and {@link #enhance(CellStyle, Workbook)}
 * methods.
 * </p>
 * 
 * @author Neil Winton
 *
 */
class CellStyleEnhancer {
	private static def workbookMap = Collections.synchronizedMap([:])
	
	/**
	 * Sets the workbook to which the CellStyle belongs. Note that this
	 * must be called before any of the other methods except {@link #enhance(CellStyle, Workbook)}.
	 * 
	 * @param self the CellStyle instance
	 * @param workbook the workbook to which it belongs
	 */
	static void setWorkbook(CellStyle self, Workbook workbook) {
		workbookMap[System.identityHashCode(self) + ":" + self.index] = workbook
	}
	
	/**
	 * Gets the workbook associated with the given CellStyle instance
	 * 
	 * @param self the CellStyle instance for which to retrieve the workbook
	 * @return the Workbook instance
	 */
	static Workbook getWorkbook(CellStyle self) {
		workbookMap[System.identityHashCode(self) + ":" + self.index]
	}
	
	
	/**
	 * Sets the data format string (as retrieved by {@link CellStyle#getDataFormatString()}
	 * for the given CellStyle instance
	 * 
	 * @param self the CellStyle instance for which to set the format string
	 * @param formatString the data format string 
	 */
	static void setDataFormatString(CellStyle self, String formatString) {
		DataFormat fmt = self.workbook.creationHelper.createDataFormat()
		self.setDataFormat(fmt.getFormat(formatString))
	}
	
	
	/**
	 * Generates a new CellStyle by combining all of the non-default style attributes.
	 * Default attributes are those found in the CellStyle at index 0 within the workbook.
	 * 
	 * @param self the first CellStyle instance
	 * @param others an array of zero or more other CellStyles to be combined with the first instance
	 * @return a new CellStyle instance with the combined styles
	 */
	static CellStyle combine(CellStyle self, CellStyle... others) {
		CellStyle combined = combineWithoutEnhancements(self, others)
		combined.workbook = self.workbook
		combined
	}
	
	/*
	 * This does the main work of combining styles. It does not assume that the instances
	 * have had permanent meta-class modifications (hence "without enhancements").
	 */
	private static CellStyle combineWithoutEnhancements(CellStyle self, CellStyle[] others) {
		Workbook workbook = self.workbook
		CellStyle base = workbook.getCellStyleAt(0 as short)
		
		def combined = [:]
		def attributes = ['alignment',
			'borderBottom', 'borderLeft', 'borderRight', 'borderTop',
			'bottomBorderColor', 'dataFormatString',
			'fillForegroundColor', 'fillBackgroundColor', // Note: Do FG before BG
			'fillPattern', 'hidden', 'indention', 'leftBorderColor',
			'locked', 'rightBorderColor', 'rotation', // 'shrinkToFit' - not present in XSSFCellStyle?
			'topBorderColor', 'verticalAlignment', 'wrapText']
		attributes.each {
			combined[it] = self[it]
		}
		
		others.each { other ->
			attributes.each { attr ->
				if (other[attr] != base[attr]) {
					combined[attr] = other[attr]
				}
			}
		}
			
		CellStyle combinedStyle = workbook.createCellStyle()
		use (CellStyleEnhancer) {
			combinedStyle.workbook = workbook
			combined.each { attribute, value ->
				combinedStyle[attribute] = value
			}
		}
		
		Font selfFont = workbook.getFontAt(self.fontIndex)
		Font[] otherFonts = others.collect { workbook.getFontAt(it.fontIndex) }
		Font combinedFont
		use (FontEnhancer) {
			selfFont.workbook = workbook
			combinedFont = selfFont.combine(otherFonts)
			combinedStyle.font = combinedFont
		}
		
		combinedStyle
	}
	
	/**
	 * This method makes permanent modifications to a CellStyle instance's meta-class
	 * to add the workbook and writable dataFormatString properties, as well as
	 * adding the {@link #combine(CellStyle, CellStyle...)} method.
	 * 
	 * @param style the CellStyle instance to enhance
	 * @param workbook the Workbook to which the CellStyle belongs
	 * @return the original CellStyle instance
	 */
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
			
			combine = { CellStyle... others ->
				enhance(CellStyleEnhancer.combineWithoutEnhancements(delegate, others), workbook)
			}
		}
		style
	}
}
