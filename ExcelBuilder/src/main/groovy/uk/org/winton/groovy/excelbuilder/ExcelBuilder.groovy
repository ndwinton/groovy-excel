package uk.org.winton.groovy.excelbuilder

import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelBuilder extends BuilderSupport {

	final Workbook workbook
	final Map<String, Sheet> sheets = [:]
	
	private String nextSheetName = 'Sheet1'
	
	ExcelBuilder() {
		workbook = new XSSFWorkbook()
	}

	@Override
	protected void setParent(Object parent, Object child) {
		// TODO Auto-generated method stub
		
	}

	@Override
	protected Object createNode(Object name) {
		switch (name) {
			case 'sheet':
				return createSheet([name: nextSheetName++])
		}
		return null
	}

	@Override
	protected Object createNode(Object name, Object value) {
		switch (name) {
			case 'sheet':
				return createSheet([name: value])
		}
		return null
	}

	@Override
	protected Object createNode(Object name, Map attributes) {
		switch (name) {
			case 'sheet':
				return createSheet(attributes)
		}
		return null;
	}

	@Override
	protected Object createNode(Object name, Map attributes, Object value) {
		// TODO Auto-generated method stub
		return null;
	}

	private Sheet createSheet(Map attributes) {
		sheets[attributes.name] = workbook.createSheet(attributes.name)
	}
}
