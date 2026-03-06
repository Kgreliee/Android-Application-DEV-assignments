package studentGradeCalculator

import androidx.compose.runtime.getValue
import androidx.compose.runtime.setValue

import android.content.Context
import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream

// 1. The Processing Logic (Optimized for Android URIs)
class GradeCalculator {
    fun calculate(context: Context, inputUri: Uri, outputFile: File) {
        // Use ContentResolver to open the file selected by the user
        context.contentResolver.openInputStream(inputUri).use { inputStream ->
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0)

            for (row in sheet) {
                if (row.rowNum == 0) continue // Skip header
                val scoreCell = row.getCell(1)
                if (scoreCell != null) {
                    val score = try { scoreCell.numericCellValue } catch (e: Exception) { 0.0 }
                    val grade = when {
                        score >= 90 -> "A"
                        score >= 80 -> "B"
                        score >= 70 -> "C"
                        score >= 60 -> "D"
                        else -> "F"
                    }
                    val gradeCell = row.createCell(2)
                    gradeCell.setCellValue(grade)
                }
            }

            // Save the result to the provided file path
            FileOutputStream(outputFile).use { outputStream ->
                workbook.write(outputStream)
            }
            workbook.close()
        }
    }
}

// 2. The Android Activity
class MainApp : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            MaterialTheme {
                Surface(modifier = Modifier.fillMaxSize(), color = MaterialTheme.colorScheme.background) {
                    GradeCalculatorScreen()
                }
            }
        }
    }
}

// 3. The UI (Fixed with File Picker)
@Composable
fun GradeCalculatorScreen() {
    val context = LocalContext.current
    val calculator = remember { GradeCalculator() }
    var statusText by remember { mutableStateOf("Tap the button to select an Excel file") }
    var isLoading by remember { mutableStateOf(false) }

    // Launcher for the Android File Picker
    val filePickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        if (uri != null) {
            isLoading = true
            try {
                // We save the processed file to the app's internal "files" directory
                val outputFile = File(context.filesDir, "Graded_Results.xlsx")

                calculator.calculate(context, uri, outputFile)

                statusText = "Success! File saved to internal storage as: ${outputFile.name}"
            } catch (e: Exception) {
                statusText = "Error: ${e.localizedMessage}"
            }
            isLoading = false
        }
    }

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(24.dp),
        horizontalAlignment = Alignment.CenterHorizontally,
        verticalArrangement = Arrangement.Center
    ) {
        Text(text = "Student Grade Calculator", style = MaterialTheme.typography.headlineMedium)

        Spacer(modifier = Modifier.height(20.dp))

        Text(
            text = "Select an .xlsx file. The app will read scores from Column B and write grades into Column C.",
            style = MaterialTheme.typography.bodySmall
        )

        Spacer(modifier = Modifier.height(30.dp))

        if (isLoading) {
            CircularProgressIndicator()
        } else {
            Button(
                onClick = {
                    // Open the picker for Excel files
                    filePickerLauncher.launch("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                },
                modifier = Modifier.fillMaxWidth(),
                colors = ButtonDefaults.buttonColors(
                    containerColor = Color(0xFFB6E7C9),
                    contentColor = Color.Black
                )
            ) {
                Text("SELECT & PROCESS EXCEL")
            }
        }

        Spacer(modifier = Modifier.height(20.dp))

        Text(
            text = statusText,
            color = if (statusText.contains("Error")) Color.Red else Color.DarkGray,
            style = MaterialTheme.typography.bodyMedium
        )
    }
}