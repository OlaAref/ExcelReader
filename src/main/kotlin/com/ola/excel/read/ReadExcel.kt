package com.olaaref.com.ola.excel.read

import com.olaaref.com.ola.excel.model.EStatement
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream

class ReadExcel {

    fun readEstatementExcel(filePath: String): List<EStatement> {
        var statements = mutableListOf<EStatement>()
        FileInputStream(File(filePath)).use { inputStream ->
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0)
            val rowIterator = sheet.rowIterator()
            while (rowIterator.hasNext()){
                val row: Row = rowIterator.next()
                val id = row.getCell(0).numericCellValue.toInt()
                val subtitleAr = row.getCell(1).stringCellValue
                val subtitleEn = row.getCell(2).stringCellValue
                val isFreeValue = row.getCell(3).stringCellValue
                val isFree = isFreeValue == "YES"
                val statement = EStatement(id, subtitleAr, subtitleEn, isFree)
                println(generateSql(statement ))
                statements.add(statement)
            }
//            workbook.close()
        }
        return statements
    }

    fun generateSql(statement: EStatement): String {
        return "INSERT INTO ESTATEMENT VALUES (${statement.id}, '${statement.subtitleAr}', '${statement.subtitleEn}', ${statement.isFree});"
    }
}