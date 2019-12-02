Imports DocumentFormat.OpenXml
Imports ClosedXML.Excel
Imports System.IO
Imports System.Text

Module Datenumwandler

    Sub Main()

        'Abfragen und Öffnen der zu umwandelnden Excel-Datei
        Console.WriteLine("Geben Sie den Namen der Datei ein, die umgewandelt werden soll")
        Dim name As String = Console.ReadLine()
        Dim excelFile As Object = OpenExcel(name)

        'Erstellen der neuen Excel-Datei
        Dim newWorkbook = New XLWorkbook()
        Dim newWorksheet As ClosedXML.Excel.IXLWorksheet = newWorkbook.Worksheets.Add("Results" & Date.Now.ToShortDateString)

        'Anlegen einer Liste, in der alle Fragebögen gespeichert werden
        Dim listOfQuestionnaires As New List(Of Questionnaire)
        Dim y As Integer = 2
        Dim cell = excelFile.application.cells(y, 1).value

        'Hier werden jetzt alle Zeilen der Datei, die umgewanldet werden soll, durchgegangen. Jede Zeile enthält den Inhalt eines Fragebogens. 
        While Not IsNothing(cell)

            'Für jeden Fragebogen wird eine neue Instanz der Klasse Questionnaire angelegt, die alle nötigen Informationen speichert
            Dim newQuestionnaire As New Questionnaire(y, excelFile)

            'Damit in der neue Excel-Datei nicht mehrmals der identische Schulcode mit identischer laufender Nummer zu finden ist,
            'wird hier abgefragt, ob der Schulcode mit der laufenden Nummer bereits abgespeichert ist.
            If Not listOfQuestionnaires.Any(Function(x) x.id = newQuestionnaire.id) And Not newQuestionnaire.id = "" Then
                listOfQuestionnaires.Add(newQuestionnaire)
            End If

            y = y + 1
            cell = excelFile.application.cells(y, 1).value

        End While

        'Hier werden alle Fragebögen sortiert anhand der Anzahl der gestellten Fragen, sodass in der neuen Excel-Datei oben die Fragebögen 
        'mit den meisten Fragen zu finden sind
        listOfQuestionnaires = listOfQuestionnaires.OrderByDescending(Function(x) x.qcount).ToList()

        'Hier werden nun die sortierten und überprüften Fragebögen in die neue Excel-Datei geschrieben
        Dim i As Integer = 1
        For Each questionnaire In listOfQuestionnaires
            questionnaire.writoToExcel(newWorksheet, i)
            i = i + 1
        Next

        'Schließen der Auslese Datei und speichern und öffnen der neuen Excel Datei
        SaveAndClose(excelFile, newWorkbook, name)

    End Sub

    Private Function OpenExcel(ByVal name As String) As Object

        'Festlegen des Pfades. Die Datei die umgewandelt werden soll, muss sich um gleichen Ordner befinden, wie diese Anwendung
        Dim excelPath As String = My.Application.Info.DirectoryPath & "\" & name & ".xlsx"

        'Auslesen der Excel-Datei
        Dim excelFile As Object = CreateObject("Excel.Application")
        excelFile.application.workbooks.Open(excelPath)
        excelFile.application.sheets(1).Select()

        Return excelFile

    End Function

    Private Sub SaveAndClose(ByVal excelFile As Object, ByVal newWorkbook As XLWorkbook, ByVal name As String)

        'Schließen der Auslese Datei
        excelFile.application.DisplayAlerts = False
        excelFile.ActiveWorkbook.Close()
        excelFile.application.DisplayAlerts = True
        excelFile.application.Quit()
        excelFile = Nothing

        'Speichern und öffnen der neuen Excel Datei
        Dim filename As String = My.Application.Info.DirectoryPath & "\" & name & "_umgewandeltsortiert.xlsx"
        newWorkbook.SaveAs(filename)
        Process.Start(filename)

    End Sub

End Module