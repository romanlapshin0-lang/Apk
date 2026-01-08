package com.example.missingnumbers

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.Button
import android.widget.TextView
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class MainActivity : AppCompatActivity() {

    private lateinit var tvResult: TextView

    private val pickExcelLauncher =
        registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri: Uri? ->
            if (uri == null) return@registerForActivityResult

            // Persist permission so the file remains accessible
            try {
                contentResolver.takePersistableUriPermission(
                    uri,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION
                )
            } catch (_: SecurityException) {
                // Some providers don't allow persistable permission; ignore.
            }

            try {
                val numbers = readNumbersFromXlsx(uri)
                val missing = findMissingNumbers(numbers)

                tvResult.text = buildString {
                    append("Найдено номеров: ${numbers.size}\n")
                    if (numbers.isNotEmpty()) {
                        append("Диапазон: ${numbers.minOrNull()} .. ${numbers.maxOrNull()}\n\n")
                    } else {
                        append("Диапазон: —\n\n")
                    }

                    if (missing.isEmpty()) {
                        append("Пропусков нет ✅")
                    } else {
                        append("Пропущенные номера (${missing.size}):\n")
                        append(missing.joinToString(", "))
                    }
                }
            } catch (e: Exception) {
                tvResult.text = "Ошибка чтения файла: ${e.message}"
            }
        }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        tvResult = findViewById(R.id.tvResult)
        val btnPick: Button = findViewById(R.id.btnPick)

        btnPick.setOnClickListener {
            pickExcelLauncher.launch(
                arrayOf(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            )
        }
    }

    /**
     * Reads integers from the first column (A) of the first sheet.
     * Accepts numeric cells and string cells like "12".
     */
    private fun readNumbersFromXlsx(uri: Uri): List<Int> {
        contentResolver.openInputStream(uri).use { input ->
            requireNotNull(input) { "Не удалось открыть файл" }

            XSSFWorkbook(input).use { workbook ->
                val sheet = workbook.getSheetAt(0)
                val result = mutableListOf<Int>()

                for (row in sheet) {
                    val cell = row.getCell(0) ?: continue // column A

                    val value: Int? = when (cell.cellType) {
                        CellType.NUMERIC -> cell.numericCellValue.toInt()
                        CellType.STRING -> cell.stringCellValue.trim().toIntOrNull()
                        CellType.FORMULA -> {
                            try {
                                cell.numericCellValue.toInt()
                            } catch (_: Exception) {
                                cell.stringCellValue.trim().toIntOrNull()
                            }
                        }
                        else -> null
                    }

                    if (value != null) result.add(value)
                }

                return result.distinct().sorted()
            }
        }
    }

    /**
     * Finds missing numbers inside [min..max].
     * Example: [1,2,4,7] -> [3,5,6]
     */
    private fun findMissingNumbers(sortedNumbers: List<Int>): List<Int> {
        if (sortedNumbers.isEmpty()) return emptyList()

        val min = sortedNumbers.first()
        val max = sortedNumbers.last()

        val set = sortedNumbers.toHashSet()
        val missing = ArrayList<Int>()

        for (n in min..max) {
            if (!set.contains(n)) missing.add(n)
        }
        return missing
    }
}
