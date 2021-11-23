package com.nithin.excelsheetwriter.utils

import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook


object SheetUtils {

    private var sheet: Sheet? = null
    const val EXCEL_SHEET_NAME = "SheetNameHere"
    private const val EMPTY = "-"
    private val arrayOfRows: ArrayList<Row> = ArrayList()

    //create workbook
    private val workbook = createWorkBookAndSheet(EXCEL_SHEET_NAME)

    private fun createWorkBookAndSheet(sheetNameHere: String): Workbook {
        val workbook = HSSFWorkbook()
        sheet = workbook.createSheet(sheetNameHere)
        return workbook
    }

    private fun addRowToSheet(rowNo: Int): Row? {
        return sheet?.createRow(rowNo)
    }

    fun setupSheet(noOfRows: Int = 2) {
        for (i in 0..noOfRows) {
            addRowToSheet(rowNo = i)/*?.let {
                arrayOfRows.add(it)
            }*/
        }
    }

    /**
     * change style based on requirement for columns or rows
     */
    private fun Workbook.setupCellStyle(): CellStyle? {
        val cellStyle = createCellStyle()
        cellStyle.fillForegroundColor = HSSFColor.AQUA.index;
        cellStyle.fillPattern = HSSFCellStyle.SOLID_FOREGROUND;
        cellStyle.alignment = CellStyle.ALIGN_CENTER;
        return cellStyle
    }

    /**
     * row number for column is always 0,
     * defaut column number is always 0
     */
    fun addColumnName(columnNo: Int = 0, columnName: String = EMPTY) {
        setupValueForCell(rowNo = 0, cellNo = columnNo, cellValue = columnName)
    }

    fun addRowValue(rowNo: Int = 1, rowValues: ArrayList<String>) {
        rowValues.forEachIndexed { index, str ->
            setupValueForCell(rowNo, cellNo = index, cellValue = str)
        }
    }

    private fun setupValueForCell(rowNo: Int = 0, cellNo: Int = 0, cellValue: String = EMPTY) {
        val cell = sheet?.getRow(rowNo)?.createCell(cellNo)
        cell?.cellStyle = workbook.setupCellStyle()
        cell?.setCellValue(cellValue)
    }

}