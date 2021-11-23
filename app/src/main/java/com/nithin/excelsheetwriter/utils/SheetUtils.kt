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
    private const val EXCEL_SHEET_NAME = "SheetNameHere"

    private fun createWorkBook(workBookName: String): Workbook {
        return HSSFWorkbook()
    }

    fun addRowToSheet(rowNo: Int): Row? {
        return sheet?.createRow(rowNo)
    }

    fun setupSheet() {
        //create workbook
        val workbook = createWorkBook(EXCEL_SHEET_NAME)
        val row0 = addRowToSheet(rowNo = 0)
        val cellStyle = workbook.setupCellStyle()
        // Creating a cell and assigning it to a row
        val cell1 = row0?.createCell(0);
        // Setting Value and Style to the cell
        cell1?.setCellValue("First Cell")
        cell1?.cellStyle = cellStyle
    }

    private fun Workbook.setupCellStyle(): CellStyle? {
        val cellStyle = createCellStyle()
        cellStyle.fillForegroundColor = HSSFColor.AQUA.index;
        cellStyle.fillPattern = HSSFCellStyle.SOLID_FOREGROUND;
        cellStyle.alignment = CellStyle.ALIGN_CENTER;
        return cellStyle
    }

    fun addTitleColumns(columnNames: ArrayList<String>) {

    }

}