package com.sushencev.kotlinpoi

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Paths
import java.util.*
import kotlin.math.max


class ExcelAdapter private constructor() {
    companion object {
        fun open(fileName: String): Workbook {
            return WorkbookFactory.create(FileInputStream(Paths.get(fileName).toFile()))
        }

        fun write(workbook: Workbook, fileName: String) {
            val outputPath = Paths.get(fileName)

            while (true) {
                try {
                    Files.newOutputStream(outputPath).use {
                        workbook.write(it)
                    }
                    break
                } catch (e: IOException) {
                    e.printStackTrace()
                    System.err.println("Please close file and press [Enter]")
                    readLine()
                }
            }
        }
    }
}

operator fun Sheet.get(n: Int): Row {
    return this.getRow(n) ?: this.createRow(n)
}

operator fun Row.get(n: Int): Cell {
    return this.getCell(n) ?: this.createCell(n, Cell.CELL_TYPE_BLANK)
}

operator fun Row.get(columnName: String): Cell {
    val column = sheet[0].let { header ->
        (0..lastCellNum + 20).find { i ->
            // +20 due to bug in Apache POI
            try {
                header[i].stringCellValue == columnName
            } catch (e: Exception) {
                false
            }
        }
                ?: error("Не найден столбец \"$columnName\" в листе \"${sheet.sheetName}\"."
                        + " Столбцы ${(firstCellNum..lastCellNum).joinToString { "\"${header[it].stringCellValue}\"" }}")
    }
    return get(column)
}

operator fun Sheet.get(x: Int, y: Int): Cell {
    return this[y][x]
}

fun Sheet.getRows(startingFrom: Int): List<Row> {
    val lastRow = if (lastRowNum > 500000) {
        System.err.println("Warning: lastRowNum is more than 500000")
        ((0 until 100000).find { get(it).let { row -> (0..9).all { row[it].myStringCellValue.isEmpty() } } } ?: 0) - 1
    } else {
        (lastRowNum downTo startingFrom).find { get(it).let { row -> !(0..9).all { row[it].myStringCellValue.isEmpty() } } } ?: 0
    }
    return getRows(startingFrom, lastRow)
}

fun Sheet.getRows(startingFrom: Int, toIncluding: Int): List<Row> {
    return (firstRowNum + startingFrom..toIncluding).map { get(it) }
}

private fun Cell.getCellStyleOrCreateNew(): CellStyle {
    val wb = sheet.workbook
    val res = if (cellStyle == wb.getCellStyleAt(0)) wb.createCellStyle() else cellStyle
    this.cellStyle = res
    return res
}

fun Cell.setRounding(digits: Int) {
    val wb = sheet.workbook
    getCellStyleOrCreateNew().dataFormat = wb.creationHelper.createDataFormat().getFormat("0.${"0".repeat(digits)}")
}

enum class HorizontalAlignment(val constantValue: Short) {
    LEFT(CellStyle.ALIGN_LEFT),
    CENTER(CellStyle.ALIGN_CENTER),
    RIGHT(CellStyle.ALIGN_RIGHT)
}

enum class VerticalAlignment(val constantValue: Short) {
    TOP(CellStyle.VERTICAL_TOP),
    CENTER(CellStyle.VERTICAL_CENTER),
    BOTTOM(CellStyle.VERTICAL_BOTTOM)
}

fun Cell.setHorizontalAlignment(alignment: HorizontalAlignment) {
    getCellStyleOrCreateNew().alignment = alignment.constantValue
}

fun Cell.setVerticalAlignment(alignment: VerticalAlignment) {
    getCellStyleOrCreateNew().verticalAlignment = alignment.constantValue
}

private fun setCellValue(cell: Cell, value: Any) {
    when (value) {
        is Pair<*, *> -> {
            cell.setCellValue(value.first as Double)
            cell.setRounding(value.second as Int)
        }
        is String -> cell.setCellValue(value)
        is Int -> setCellValue(cell, value.toDouble())
        is Long -> setCellValue(cell, value.toDouble())// to 0
        is Double -> cell.setCellValue(value)
        is Date -> {
            // TODO: make single date style or group styles in order to reduce their amount
            cell.setCellValue(value)
            val cellStyle = cell.sheet.workbook.createCellStyle()
            cellStyle.dataFormat = 14
            cell.cellStyle = cellStyle
        }
        else -> throw IllegalArgumentException(value::class.java.simpleName)
    }
}

operator fun Row.set(field: String, value: Any) {
    setCellValue(get(field), value)
}

operator fun Row.set(y: Int, value: Any) {
    setCellValue(get(y), value)
}

fun Sheet.setRow(rowNum: Int, values: Collection<Any>) {
    if (values.size == 1 && values.single() is Row) {
        replaceRow(rowNum, values.single() as Row)
    } else {
        values.forEachIndexed { i, value -> this[rowNum][i] = value }
    }
}

fun Sheet.addRow(values: Collection<Any>) {
    setRow(lastRowNum + 1, values)
}

fun Sheet.addRow(vararg values: Any) {
    setRow(lastRowNum + 1, values.toList())
}

fun Sheet.setHeader(values: Collection<Any>) {
    setRow(0, values)
}

fun Sheet.setHeader(vararg values: Any) {
    setHeader(values.toList())
}

fun Sheet.replaceRow(rowNum: Int, newRow: Row) {
    val sourceRow = get(rowNum)
    (0..max(sourceRow.lastCellNum.toInt(), newRow.lastCellNum.toInt())).forEach {
        val cell = sourceRow[it]
        val sourceCell = newRow[it]

//        val newStyle = workbook.createCellStyle()
//        newStyle.cloneStyleFrom(sourceCell.cellStyle)
//        cell.cellStyle = newStyle
        // TODO: it is wrong, because 'cell.cellStyle' may return 'default' cell style and we don't want to modify it
        // see getCellStyleOrCreateNew() function
        cell.cellStyle.dataFormat = sourceCell.cellStyle.dataFormat

        cell.cellType = sourceCell.cellType
        setCellValue(cell, try {
            sourceCell.stringCellValue
        } catch (e: Throwable) {
            sourceCell.numericCellValue
        })
        setColumnWidth(it, newRow.sheet.getColumnWidth(it))
    }
}

val Sheet.lastRow: Row
    get() = get(lastRowNum)

fun Cell.isNumericValue(): Boolean {
    return try {
        numericCellValue
        true
    } catch (e: Exception) {
        false
    }
}

val Cell.myStringCellValue: String
    get() = try {
        stringCellValue
    } catch (e: Exception) {
        try {
            numericCellValue.toInt().toString()
        } catch (e1: Exception) {
            System.err.println(e)
            throw e1
        }
    }

val Cell.myDoubleCellValue: Double
    get() = try {
        numericCellValue
    } catch (e: Exception) {
        stringCellValue.trim().toDouble()
    }

val Cell.myLongCellValue: Long
    get() = myDoubleCellValue.toLong()

val Cell.preferablyStringValue: Any
    get() = try {
        stringCellValue
    } catch (e: Exception) {
        numericCellValue.toInt()
    }

val Cell.value: Any
    get() = try {
        stringCellValue
    } catch (e: Exception) {
        numericCellValue
    }

val Cell.preferablyLongValue: String
    get() = try {
        numericCellValue.toLong().toString()
    } catch (e: Exception) {
        stringCellValue
    }

val Cell.preferablyDoubleValue: Any
    get() = try {
        numericCellValue
    } catch (e: Exception) {
        stringCellValue
    }


fun Cell.createDropdownList(hiddenName: String, choices: List<String>) {
    val validationHelper = XSSFDataValidationHelper(sheet as XSSFSheet)
    val addressList = CellRangeAddressList(rowIndex, rowIndex, columnIndex, columnIndex)

    val hiddenSheet = sheet.workbook.createSheet(hiddenName)
    sheet.workbook.setSheetHidden(hiddenSheet.workbook.getSheetIndex(hiddenSheet), 2)
    choices.forEach { hiddenSheet.addRow(it) }
    val namedCell = sheet.workbook.createName()
    namedCell.nameName = hiddenName
    namedCell.refersToFormula = "$hiddenName!\$A$1:\$A\$${choices.size + 1}"

    val constraint = validationHelper.createFormulaListConstraint(hiddenName)
    val dataValidation = validationHelper.createValidation(constraint, addressList)
    dataValidation.suppressDropDownArrow = true
    sheet.addValidationData(dataValidation)
}

fun Workbook.createSheetWithHeader(sheetName: String, header: List<String>, autoSizeColumns: Boolean = true, runnable: (sheet: Sheet) -> Unit) {
    val sheet = createSheet(sheetName)
    sheet.setHeader(header)
    runnable(sheet)
    if (autoSizeColumns) (0 until header.size).forEach { sheet.autoSizeColumn(it) }
}

fun Workbook.renameSheet(oldName: String, newName: String) {
//    setSheetName(getSheetIndex(oldName), newName)
    (this as XSSFWorkbook).ctWorkbook.sheets.getSheetArray(getSheetIndex(oldName)).name = newName
}

fun setRegionBorderWithMedium(region: CellRangeAddress, sheet: Sheet) {
    val wb = sheet.workbook
    RegionUtil.setBorderBottom(CellStyle.BORDER_MEDIUM.toInt(), region, sheet, wb)
    RegionUtil.setBorderLeft(CellStyle.BORDER_MEDIUM.toInt(), region, sheet, wb)
    RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM.toInt(), region, sheet, wb)
    RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM.toInt(), region, sheet, wb)
}