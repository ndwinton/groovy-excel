# ExcelBuilder

This provides a builder-style interface for creating Excel (2007+ format) spreadsheets using the Apache POI library.

## Quick Start

To give you an idea of the main features here is a quick example:

    def builder = new ExcelBuilder()
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
    builder.workbook.write(new FileOutputStream(new File('test.xlsx')))

This produces a workbook with a sheet that looks something like this:

![Screenshot of resulting spreadsheet](test-screenshot.png "Screenshot")

## Leisurely Instructions

The ExcelBuilder class provides a simple way to create an Excel (Open Office XML) format
spreadsheet. It follows the model of other "builder"-style implementations such as
the [MarkupBuilder](http://groovy.codehaus.org/Creating+XML+using+Groovy's+MarkupBuilder).

### Creating the Builder

You can create a new builder instance, representing a completely empty workbook,
using the default constructor:

    def builder = new ExcelBuilder()

This builder will contain no worksheets or any other content. Alternatively, you can use an
existing workbook file as a template by supplying a File instance or filename as a string:

    def builderFromTemplate = new ExcelBuilder("template.xlsx")

In the latter case you can add to or modify the existing contents of the template workbook
(although the template file itself won't be altered unless you explicitly overwrite it).

### Adding Sheets

A workbook contains one or more sheets of data. You can add a sheet to the workbook by
using the builder as follows:

    builder {
        sheet {
            // Sheet contents here ...
        }
    }
 
By default the first sheet added will be called "Sheet1", the next "Sheet2" and so on.
You can explicitly name the sheet by supplying the name as a parameter, like this:
 
    builder {
        sheet("My Sheet") {
            // Sheet contents here ...
        }
    }
 
This will, unsurprisingly, create a sheet called "My Sheet". If you created the builder from
a template file then modifications will be made to the sheet with the given name, if it already
exists. Such sheets should, preferably, be identified by an explicit name, but it also works
with the implicit, automatically generated names.

The _sheet()_ method within the builder returns an _org.apache.poi.ss.usermodel.Sheet_ instance.
This means that you can invoke further operations on the sheet which are not directly
supported by the builder. For example:

    Sheet s
    builder { s = sheet("First") }
    assert s.sheetName == 'First'

Note that you could also get the same result as follows:

    Sheet s = builder.sheet("First")
    assert s.sheetName == 'First'

Once a sheet has been created, it's _Sheet_ instance can also be accessed via the _sheets_
property of the builder:

    builder {
        sheet()
        sheet('Second')
    }
    assert builder.sheets['Sheet1'].sheetName == 'Sheet1'
    assert builder.sheets['Second'].sheetName == 'Second'

Sheets loaded from a can also be accessed in this manner.

### Fonts and Styles

### Creating Rows

The simplest way to add data to a worksheet is to use the _row()_ method. Each row can
take a list that describes the contents of the row (starting from the first cell within
the row).

    builder {
        sheet("My Sheet") {
            row([1, 2, 3.14159, 'Hello, world!', new Date()])
        }
    }

This will create a single row within the "My Sheet" worksheet containing three numeric
cells, a string and a date. Unless otherwise specified, rows will be created starting
from the first row in the sheet (numbered 0). The row number can be explicitly set using
the _row_ attribute. So, the following code will generate alternating rows of content:

    builder {
        sheet("My Sheet") {
            for (r = 0; r < 10; r += 2) {
                row([r, "This is row " + r], row: r)
            }
        }
    }
    
Any row added after an explicitly positioned row will be added at the next position.

The _row()_ method within the builder returns an _org.apache.poi.ss.usermodel.Row_ instance.
As with sheets, you can use the normal Apache POI methods to operate on this object.

#### Height, Width and Styles

## The "Enhancer" Classes

There are several "enhancer" classes, used within the builder which add functionality
to the base Apache POI classes. They can also be used separately as category classes
(e.g. as `use(FontEnhancer) { ... }`). Each of the classes has full GroovyDoc that
describes them, but briefly the functionality added by each of the classes is
as listed below.

### CellStyleEnhancer

* _workbook_ property  
* _dataFormatString_ property  
* _combine(CellStyle self, CellStyle... others)_ method  
* _enhance(CellStyle target)_ method

### FontEnhancer

* _workbook_ property
* _bold_ property
* _fontHeightInPoints_ property
* _combine(Font self, Font... others)_ method
* _enhance(Font target)_ method

### SheetEnhancer

* _active_ property
* _hidden_ property
* _veryHidden_ property
* _defaultColumnWidthInChars_ property
* _getColumnWidthInChars(int columnIndex)_ method
* _setColumnWidthInChars(int columnIndex, Number width)_ method


