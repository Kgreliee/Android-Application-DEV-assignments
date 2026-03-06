package studentGradeCalculator

sealed class Grade(val letter: String) {
    object A : Grade("A")
    object B : Grade("B")
    object C : Grade("C")
    object F : Grade("F")
}
