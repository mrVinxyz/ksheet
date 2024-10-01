import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.time.Instant
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.time.format.DateTimeFormatter
import kotlin.io.path.Path
import kotlin.io.use
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertNotNull
import kotlin.test.assertTrue

class SheetTest {
    @Test
    fun `should write and verify excel file content`() {
        data class Item(
            val strField: String,
            val intField: Int,
            val longField: Int,
            val floatField: Float,
            val doubleField: Double,
            val boolField: Boolean,
            val dateFieldString: String,
            val dateFieldLocalDate: LocalDate,
            val dateFieldLocalDateTime: LocalDateTime,
            val dateFieldInstant: Instant
        )

        val formatter = DateTimeFormatter.ISO_LOCAL_DATE

        val itemList = listOf(
            Item(
                "John Doe", 30, 1000, 1.5f, 2.5, true,
                "2024-01-01",
                LocalDate.parse("2024-01-01", formatter),
                LocalDateTime.parse("2024-01-01T00:00:00"),
                Instant.parse("2024-01-01T00:00:00Z")
            ),
            Item(
                "Jane Smith", 25, 2000, 2.5f, 3.5, false,
                "2023-01-01",
                LocalDate.parse("2023-01-01", formatter),
                LocalDateTime.parse("2023-01-01T00:00:00"),
                Instant.parse("2023-01-01T00:00:00Z")
            ),
            Item(
                "Alice Johnson", 35, 3000, 3.5f, 4.5, true,
                "2022-01-01",
                LocalDate.parse("2022-01-01", formatter),
                LocalDateTime.parse("2022-01-01T00:00:00"),
                Instant.parse("2022-01-01T00:00:00Z")
            )
        )

        val sheetPath = Path("example.xlsx")

        val sheet = workbook {
            sheet("Sheet 1") {
                row {
                    cell("Name")
                    cell("Age")
                    cell("Long Field")
                    cell("Float Field")
                    cell("Double Field")
                    cell("Boolean Field")
                    cell("Date String")
                    cell("Date LocalDate")
                    cell("Date LocalDateTime")
                    cell("Date Instant")
                }

                rows(itemList) { item ->
                    cell(item.strField)
                    cell(item.intField)
                    cell(item.longField)
                    cell(item.floatField)
                    cell(item.doubleField)
                    cell(item.boolField)
                    // Adding multiple date types
                    cell(item.dateFieldString)
                    cell(item.dateFieldLocalDate)
                    cell(item.dateFieldLocalDateTime)
                    cell(item.dateFieldInstant)
                }
            }
        }

        sheet.intoFile(sheetPath)

        assert(File(sheetPath.toString()).exists())

        FileInputStream("example.xlsx").use { fis ->
            val workbook = WorkbookFactory.create(fis)
            val sheet = workbook.getSheetAt(0)
            assertNotNull(sheet)

            val headerRow = sheet.getRow(0)
            assertEquals("Name", headerRow.getCell(0).stringCellValue)
            assertEquals("Age", headerRow.getCell(1).stringCellValue)

            val firstDataRow = sheet.getRow(1)
            assertEquals("John Doe", firstDataRow.getCell(0).stringCellValue)
            assertEquals(30, firstDataRow.getCell(1).numericCellValue.toInt())
            assertEquals(1000L, firstDataRow.getCell(2).numericCellValue.toLong())
            assertEquals(1.5f, firstDataRow.getCell(3).numericCellValue.toFloat())
            assertEquals(2.5, firstDataRow.getCell(4).numericCellValue)
            assertEquals(true, firstDataRow.getCell(5).booleanCellValue)
            assertEquals("2024-01-01", firstDataRow.getCell(6).stringCellValue)

//            val localDateCell = firstDataRow.getCell(7)
//            assertTrue(DateUtil.isCellDateFormatted(localDateCell))
//            val expectedLocalDate = LocalDate.parse("2024-01-01", formatter)
//            val actualLocalDate = localDateCell.localDateTimeCellValue.toLocalDate()
//            assertEquals(expectedLocalDate, actualLocalDate)
//
//            val localDateTimeCell = firstDataRow.getCell(8)
//            assertTrue(DateUtil.isCellDateFormatted(localDateTimeCell))
//            val expectedLocalDateTime = LocalDateTime.parse("2024-01-01T00:00:00")
//            val actualLocalDateTime = localDateTimeCell.localDateTimeCellValue
//            assertEquals(expectedLocalDateTime, actualLocalDateTime)
//
//            val instantCell = firstDataRow.getCell(9)
//            assertTrue(DateUtil.isCellDateFormatted(instantCell))
//            val expectedInstant = Instant.parse("2024-01-01T00:00:00Z")
//            val actualInstant = instantCell.localDateTimeCellValue.atZone(ZoneId.systemDefault()).toInstant()
//            assertEquals(expectedInstant, actualInstant)
        }
    }
}