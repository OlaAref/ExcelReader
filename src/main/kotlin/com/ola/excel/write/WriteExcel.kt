package com.olaaref.com.ola.excel.write

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

class WriteExcel {
    fun writeExcel(filePath: String) {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Sheet1")

        val row = sheet.createRow(0)
        val cell = row.createCell(0)
        cell.setCellValue("Hello, Excel!")

        FileOutputStream(filePath).use { outputStream ->
            workbook.write(outputStream)
        }
        workbook.close()
        println("Excel file written successfully.")
    }
}