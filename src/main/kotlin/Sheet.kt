import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.xssf.usermodel.XSSFCreationHelper as XSLXCellStyleHelper
import org.apache.poi.xssf.usermodel.XSSFCellStyle as XSLXCellStyle
import org.apache.poi.ss.usermodel.Row as XSLXRow
import org.apache.poi.ss.usermodel.Sheet as XSLXSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook as XSLXWorkbook
import java.io.FileOutputStream
import java.nio.file.Path
import java.time.Instant
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.Calendar
import java.util.Date
import kotlin.apply
import kotlin.collections.forEach
import kotlin.io.use
import kotlin.let
import kotlin.toString

class Workbook {
    private val workbook: XSLXWorkbook = XSLXWorkbook()
    private var sheetIndex: Int = 0

    fun sheet(name: String?, sheetBlock: Sheet.() -> Unit) {
        val sheet = workbook.createSheet(name ?: "Sheet $sheetIndex")
        Sheet(workbook, sheet).apply(sheetBlock)
        sheetIndex++
    }

    fun intoFile(filePath: Path) {
        FileOutputStream(filePath.toString()).use { fileOut ->
            workbook.write(fileOut)
        }
    }
}

class Sheet(private val workbook: XSLXWorkbook, private val sheet: XSLXSheet) {
    private var rowIndex = 0

    fun row(rowBlock: Row.() -> Unit) {
        val row = sheet.createRow(rowIndex++)
        Row(workbook, row).apply(rowBlock)
    }

    fun <T> rows(values: List<T>, rowBlock: Row.(T) -> Unit) {
        values.forEach { value -> row { rowBlock(value) } }
    }
}

class Row(private val workbook: XSLXWorkbook, private val row: XSLXRow) {
    private var cellIndex = 0

    fun cell(value: Any? = null, styleBlock: ((CellStyler) -> Unit)? = null) {
        val cell = Cell(workbook, row, cellIndex++, value)
        cell.applyStyle(styleBlock)
    }
}

class Cell(
    private val workbook: XSLXWorkbook,
    row: XSLXRow,
    index: Int,
    value: Any? = null,
) {
    val cell = row.createCell(index)

    init {
        when (value) {
            is String -> cell.setCellValue(value)
            is Int -> cell.setCellValue(value.toDouble())
            is Long -> cell.setCellValue(value.toDouble())
            is Float -> cell.setCellValue(value.toDouble())
            is Double -> cell.setCellValue(value)
            is Boolean -> cell.setCellValue(value)
            is Calendar -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is LocalDate -> cell.setCellValue(Date.from(value.atStartOfDay(ZoneId.systemDefault()).toInstant()))
            is LocalDateTime -> cell.setCellValue(Date.from(value.atZone(ZoneId.systemDefault()).toInstant()))
            is Instant -> cell.setCellValue(Date.from(value))
            else -> cell.setCellValue(value.toString())
        }
    }

    fun applyStyle(styleBlock: ((CellStyler) -> Unit)? = null) {
        styleBlock?.let { block ->
            val cellStyle = workbook.createCellStyle()
            block(CellStyler(cellStyle, workbook.creationHelper))
            cell.cellStyle = cellStyle
        }
    }
}

class CellStyler(
    private val style: XSLXCellStyle,
    private val helper: XSLXCellStyleHelper,
) : CellStyle by style, CreationHelper by helper

fun workbook(block: Workbook.() -> Unit): Workbook = Workbook().apply(block)
