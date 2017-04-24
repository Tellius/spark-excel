package com.crealytics.spark.excel

import java.text.SimpleDateFormat

import org.apache.poi.ss.usermodel.{Cell, DataFormatter, DateUtil}

object CustomDataFormatter {

  val defaultDateFormatString = "yyyy-mm-dd hh:mm:ss"

  def formatCellValue(cell: Cell): String = {
    val _formatter = new DataFormatter()
    cell.getCellType match {
      case Cell.CELL_TYPE_NUMERIC =>
        val style = cell.getCellStyle
        if (style == null)
          cell.getNumericCellValue.toString
        else if(DateUtil.isADateFormat(style.getDataFormat, style.getDataFormatString)){
          val newFormat = cell.getSheet.getWorkbook.createDataFormat()
          val formatIndex = newFormat.getFormat(defaultDateFormatString)
          _formatter.formatRawCellContents(cell.getNumericCellValue, formatIndex, newFormat.getFormat(formatIndex))
        }
        else
          _formatter.formatRawCellContents(cell.getNumericCellValue, style.getDataFormat, style.getDataFormatString())

      case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue.toString
      case Cell.CELL_TYPE_ERROR => "ERROR"
      case Cell.CELL_TYPE_STRING => cell.getRichStringCellValue.getString
      case Cell.CELL_TYPE_FORMULA =>
        cell.setCellType(cell.getCachedFormulaResultType)
        formatCellValue(cell)
      case _ =>
        _formatter.formatCellValue(cell)
    }
  }

}
