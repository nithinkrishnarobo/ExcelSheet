package com.nithin.excelsheetwriter.utils

import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*


object SampleUtil {

    // Global Variables
    private var cell: Cell? = null
    private var sheet: Sheet? = null

    private const val EXCEL_SHEET_NAME = "Sheet1"

    /**
     * Method: Generate Excel Workbook
     */
    fun createExcelWorkbook() {
        // New Workbook
        val workbook: Workbook = HSSFWorkbook()
        // Cell style for header row
        val cellStyle = workbook.createCellStyle()
        cellStyle.fillForegroundColor = HSSFColor.AQUA.index
        cellStyle.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
        cellStyle.alignment = CellStyle.ALIGN_CENTER

        sheet = workbook.createSheet(EXCEL_SHEET_NAME)

        // Generate column headings
        val row: Row? = sheet?.createRow(0)
        cell = row?.createCell(0)
        cell?.setCellValue("First Name")
        cell?.cellStyle = cellStyle
        cell = row?.createCell(1)
        cell?.setCellValue("Last Name")
        cell?.cellStyle = cellStyle
        cell = row?.createCell(2)
        cell?.setCellValue("Phone Number")
        cell?.cellStyle = cellStyle
        cell = row?.createCell(3)
        cell?.setCellValue("Mail ID")
        cell?.cellStyle = cellStyle
    }

}