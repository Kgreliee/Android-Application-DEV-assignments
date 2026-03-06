package studentGradeCalculator

fun main() {
    val inputFile = "students.xlsx" // Path to the input Excel file
    val outputFile = "students_with_grades.xlsx" // Path to the output Excel file

    val calculator = GradeCalculator()
    calculator.calculate(inputFile, outputFile)

    println("Grades have been calculated and saved to $outputFile")
}
