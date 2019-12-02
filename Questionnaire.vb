Public Class Questionnaire

    Property schoolcode As String
    Property number As String
    Property id As String

    Property questions As New List(Of String)
    Property qcount As Integer = 0

    Property assesments As New List(Of String)
    Property gender As String

    Sub New(ByVal y As Integer, ByVal excelFile As Object)

        'In der 9. Spalte steht der Schulcode
        schoolcode = excelFile.application.cells(y, 9).value

        'In der 10. Spalte steht die laufende Nummer
        number = excelFile.application.cells(y, 10).value

        id = schoolcode & number

        'Von Spalte 11 bis 25 stehen die Fragen
        For i As Integer = 11 To 25
            Dim question = excelFile.application.cells(y, i).value
            questions.Add(question)
            If Not question = "" Then
                qcount = qcount + 1
            End If
        Next

        'Von Spalte 26 bis 30 stehen die EInschätzungen
        For i As Integer = 26 To 30
            assesments.Add(excelFile.application.cells(y, i).value)
        Next

        'In Spalte 31 steht das Geschlecht
        gender = excelFile.application.cells(y, 31).value

    End Sub

    Sub writoToExcel(ByVal excelWorksheet As ClosedXML.Excel.IXLWorksheet, ByVal i As Integer)

        'In der 1. Spalte soll die Anzahl der Fragen stehen
        excelWorksheet.Cell(i, 1).Value = qcount

        'In der 2. Spalte soll der Schulcode stehen und in der 2. Spalte die laufende Nummer
        excelWorksheet.Cell(i, 2).Value = schoolcode
        excelWorksheet.Cell(i, 3).Value = number

        'Von Spalte 4 zu 18 stehen die Fragen
        Dim k As Integer = 0
        For j As Integer = 4 To 18
            excelWorksheet.Cell(i, j).Value = questions(k)
            k = k + 1
        Next

        'Von Spalte 19 zu 23 stehen die Einschätzungen
        k = 0
        For j As Integer = 19 To 23
            excelWorksheet.Cell(i, j).Value = assesments(k)
            k = k + 1
        Next

        'In der 24. Spalte soll das Geschlecht stehen
        excelWorksheet.Cell(i, 24).Value = gender

    End Sub




End Class






