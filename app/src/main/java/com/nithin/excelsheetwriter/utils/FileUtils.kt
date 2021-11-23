package com.nithin.excelsheetwriter.utils

import android.content.Context
import android.os.Environment
import android.util.Log
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException


object FileUtils {

    private const val FILE_NAME = "Excel_File_1"

    fun createFile(context: Context, fileName: String = FILE_NAME) {
        val file = File(context.getExternalFilesDir(null), fileName)
        var fileInputStream: FileInputStream? = null
        try {
            fileInputStream = FileInputStream(file)
            Log.e("NITHIN", "Reading from Excel $file")

            // Create instance having reference to .xls file
            val workbook = HSSFWorkbook(fileInputStream)

            // Fetch sheet at position 'i' from the workbook
            val sheet = workbook.getSheetAt(0)

            // Iterate through each row
            for (row in sheet) {
                if (row.rowNum > 0) {
                    // Iterate through all the cells in a row (Excluding header row)
                    val cellIterator = row.iterator()
                    while (cellIterator.hasNext()) {
                        val cell: Cell = cellIterator.next()
                        // Check cell type and format accordingly
                        when (cell.cellType) {
                            Cell.CELL_TYPE_NUMERIC ->
                                // Print cell value
                            {
                                println(cell.numericCellValue)
                                break
                            }
                            Cell.CELL_TYPE_STRING -> {
                                println(cell.stringCellValue)
                                break
                            }
                        }
                    }
                } else {
                    Log.e("NITHIN", "Empty File")
                }
            }
        } catch (e: IOException) {
            Log.e("NITHIN", "Error Reading Exception: ", e)
        } catch (e: Exception) {
            Log.e("NITHIN", "Failed to read file due to Exception: ", e)
        } finally {
            try {
                fileInputStream?.close()
            } catch (ex: Exception) {
                ex.printStackTrace()
            }
        }
    }

    /**
     * Export Data into Excel Workbook
     *
     * @param context  - Pass the application context
     * @param fileName - Pass the desired fileName for the output excel Workbook
     * @param dataList - Contains the actual data to be displayed in excel
     */
    fun exportDataIntoWorkbook(
        context: Context, fileName: String = SheetUtils.EXCEL_SHEET_NAME,
        dataList: List<String>
    ): Boolean {

        // Check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e("NITHIN", "Storage not available or read only")
            return false
        }

        // Creating a New HSSF Workbook (.xls format)
        val workbook = HSSFWorkbook()
        val cellStyle = setHeaderCellStyle(workbook)

        // Creating a New Sheet and Setting width for each column
        val sheet = workbook.createSheet(SheetUtils.EXCEL_SHEET_NAME)
        sheet.setColumnWidth(0, 15 * 400)
        sheet.setColumnWidth(1, 15 * 400)
        sheet.setColumnWidth(2, 15 * 400)
        sheet.setColumnWidth(3, 15 * 400)
        setHeaderRow(sheet, cellStyle)
        fillDataIntoExcel(dataList, sheet)
        return storeExcelInStorage(context, "$fileName.xls", workbook)
    }

    /**
     * Checks if Storage is READ-ONLY
     *
     * @return boolean
     */
    private fun isExternalStorageReadOnly(): Boolean {
        val externalStorageState = Environment.getExternalStorageState()
        return Environment.MEDIA_MOUNTED_READ_ONLY == externalStorageState
    }

    /**
     * Checks if Storage is Available
     *
     * @return boolean
     */
    private fun isExternalStorageAvailable(): Boolean {
        val externalStorageState = Environment.getExternalStorageState()
        return Environment.MEDIA_MOUNTED == externalStorageState
    }

    /**
     * Setup header cell style
     */
    private fun setHeaderCellStyle(workbook: HSSFWorkbook): HSSFCellStyle? {
        val headerCellStyle = workbook.createCellStyle()
        headerCellStyle.fillForegroundColor = HSSFColor.AQUA.index
        headerCellStyle.fillPattern = HSSFCellStyle.SOLID_FOREGROUND
        headerCellStyle.alignment = CellStyle.ALIGN_CENTER
        return headerCellStyle
    }

    /**
     * Setup Header Row
     */
    private fun setHeaderRow(sheet: HSSFSheet, cellStyle: HSSFCellStyle?) {
        val headerRow: Row = sheet.createRow(0)
        var cell = headerRow.createCell(0)
        cell.setCellValue("First Name")
        cell.cellStyle = cellStyle
        cell = headerRow.createCell(1)
        cell.setCellValue("Last Name")
        cell.cellStyle = cellStyle
        cell = headerRow.createCell(2)
        cell.setCellValue("Phone Number")
        cell.cellStyle = cellStyle
        cell = headerRow.createCell(3)
        cell.setCellValue("Mail ID")
    }

    /**
     * Fills Data into Excel Sheet
     *
     *
     * NOTE: Set row index as i+1 since 0th index belongs to header row
     *
     * @param dataList - List containing data to be filled into excel
     */
    private fun fillDataIntoExcel(dataList: List<String>, sheet: HSSFSheet) {
        for (i in dataList.indices) {
            // Create a New Row for every new entry in list
            val rowData: Row = sheet.createRow(i + 1)

            // Create Cells for each row
            val cell1 = rowData.createCell(0)
            cell1.setCellValue(dataList[i])
            val cell2 = rowData.createCell(1)
            cell2.setCellValue(dataList[i] + "1")
            val cell3 = rowData.createCell(2)
            cell3.setCellValue(dataList[i] + "2")
            val cell4 = rowData.createCell(3)
            cell4.setCellValue(dataList[i] + "3")
        }
    }

    /**
     * Store Excel Workbook in external storage
     *
     * @param context  - application context
     * @param fileName - name of workbook which will be stored in device
     * @return boolean - returns state whether workbook is written into storage or not
     */
    private fun storeExcelInStorage(
        context: Context,
        fileName: String,
        workbook: HSSFWorkbook
    ): Boolean {
        var isSuccess: Boolean
        val file = File(context.getExternalFilesDir(null), fileName)
        var fileOutputStream: FileOutputStream? = null
        try {
            fileOutputStream = FileOutputStream(file)
            workbook.write(fileOutputStream)
            Log.e("NITHIN", "Writing file $file")
            isSuccess = true
        } catch (e: IOException) {
            Log.e("NITHIN", "Error writing Exception: ", e)
            isSuccess = false
        } catch (e: java.lang.Exception) {
            Log.e("NITHIN", "Failed to save file due to Exception: ", e)
            isSuccess = false
        } finally {
            try {
                fileOutputStream?.close()
            } catch (ex: java.lang.Exception) {
                ex.printStackTrace()
            }
        }
        return isSuccess
    }


}