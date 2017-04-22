package com.crealytics.spark.excel

import org.apache.poi.ss.usermodel.{Cell, DataFormatter}

/**
  * Created by shashank on 22/04/17.
  */
object CustomDataFormatter {

  def formatCellValue(cell:Cell):String = {
    val _formatter = new DataFormatter()
    if(cell.getCellType == Cell.CELL_TYPE_FORMULA) {
      cell.getCachedFormulaResultType match {
        case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue.toString
        case Cell.CELL_TYPE_NUMERIC =>
          val style = cell.getCellStyle
          if(style == null)
            cell.getNumericCellValue.toString
          else
            _formatter.formatRawCellContents(cell.getNumericCellValue, style.getDataFormat, style.getDataFormatString())
        case Cell.CELL_TYPE_ERROR => "ERROR"
        case Cell.CELL_TYPE_STRING => cell.getRichStringCellValue.getString
      }
    } else
      _formatter.formatCellValue(cell)
  }

}
