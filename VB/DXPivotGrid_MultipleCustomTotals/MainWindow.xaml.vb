Imports System.Windows
Imports DXPivotGrid_MultipleCustomTotals.DataSet1TableAdapters
Imports DevExpress.Xpf.PivotGrid
Imports System.Collections

Namespace DXPivotGrid_MultipleCustomTotals

    Public Partial Class MainWindow
        Inherits Window

        Private salesPersonDataAdapter As SalesPersonTableAdapter = New SalesPersonTableAdapter()

        Public Sub New()
            Me.InitializeComponent()
            ' Binds the pivot grid to data.
            Me.pivotGridControl1.DataSource = salesPersonDataAdapter.GetData()
            ' Creates a PivotGridCustomTotal object that defines the Median Custom Total.
            Dim medianCustomTotal As PivotGridCustomTotal = New PivotGridCustomTotal()
            medianCustomTotal.SummaryType = FieldSummaryType.Custom
            ' Specifies a unique PivotGridCustomTotal.Tag property value 
            ' that will be used to distinguish between two Custom Totals.
            medianCustomTotal.Tag = "Median"
            ' Specifies formatting settings that will be used to display 
            ' Custom Total column/row headers.
            medianCustomTotal.Format = "{0} Median"
            ' Adds the Median Custom Total for the Sales Person field.
            Me.fieldSalesPerson.CustomTotals.Add(medianCustomTotal)
            ' Creates a PivotGridCustomTotal object that defines the Quartiles Custom Total.
            Dim quartileCustomTotal As PivotGridCustomTotal = New PivotGridCustomTotal()
            quartileCustomTotal.SummaryType = FieldSummaryType.Custom
            ' Specifies a unique PivotGridCustomTotal.Tag property value 
            ' that will be used to distinguish between two Custom Totals.
            quartileCustomTotal.Tag = "Quartiles"
            ' Specifies formatting settings that will be used to display 
            ' Custom Total column/row headers.
            quartileCustomTotal.Format = "{0} Quartiles"
            ' Adds the Quartiles Custom Total for the Sales Person field.
            Me.fieldSalesPerson.CustomTotals.Add(quartileCustomTotal)
            ' Enables the Custom Totals to be displayed instead of Automatic Totals.
            Me.fieldSalesPerson.TotalsVisibility = FieldTotalsVisibility.CustomTotals
            Me.pivotGridControl1.RowTotalsLocation = FieldRowTotalsLocation.Far
        End Sub

        ' Handles the CustomCellValue event. 
        ' Fires for each data cell. If the processed cell is a Custom Total,
        ' provides an appropriate Custom Total value.
        Private Sub pivotGridControl1_CustomCellValue(ByVal sender As Object, ByVal e As PivotCellValueEventArgs)
            ' Exits if the processed cell does not belong to a Custom Total.
            If e.ColumnCustomTotal Is Nothing AndAlso e.RowCustomTotal Is Nothing Then Return
            ' Obtains a list of summary values against which
            ' the Custom Total will be calculated.
            Dim summaryValues As ArrayList = GetSummaryValues(e)
            ' Obtains the name of the Custom Total that should be calculated.
            Dim customTotalName As String = GetCustomTotalName(e)
            ' Calculates the Custom Total value and assigns it to the Value event parameter.
            e.Value = GetCustomTotalValue(summaryValues, customTotalName)
        End Sub

        ' Returns the Custom Total name.
        Private Function GetCustomTotalName(ByVal e As PivotCellValueEventArgs) As String
            Return If(e.ColumnCustomTotal IsNot Nothing, e.ColumnCustomTotal.Tag.ToString(), e.RowCustomTotal.Tag.ToString())
        End Function

        ' Returns a list of summary values against which
        ' a Custom Total will be calculated.
        Private Function GetSummaryValues(ByVal e As PivotCellValueEventArgs) As ArrayList
            Dim values As ArrayList = New ArrayList()
            ' Creates a summary data source.
            Dim sds As PivotSummaryDataSource = e.CreateSummaryDataSource()
            ' Iterates through summary data source records
            ' and copies summary values to an array.
            For i As Integer = 0 To sds.RowCount - 1
                Dim value As Object = sds.GetValue(i, e.DataField)
                If value Is Nothing Then
                    Continue For
                End If

                values.Add(value)
            Next

            ' Sorts summary values.
            values.Sort()
            ' Returns the summary values array.
            Return values
        End Function

        ' Returns the Custom Total value by an array of summary values.
        Private Function GetCustomTotalValue(ByVal values As ArrayList, ByVal customTotalName As String) As Object
            ' Returns a null value if the provided array is empty.
            If values.Count = 0 Then
                Return Nothing
            End If

            ' If the Median Custom Total should be calculated,
            ' calls the GetMedian method.
            If Equals(customTotalName, "Median") Then
                Return GetMedian(values)
            End If

            ' If the Quartiles Custom Total should be calculated,
            ' calls the GetQuartiles method.
            If Equals(customTotalName, "Quartiles") Then
                Return GetQuartiles(values)
            End If

            ' Otherwise, returns a null value.
            Return Nothing
        End Function

        ' Calculates a median for the specified sorted sample.
        Private Function GetMedian(ByVal values As ArrayList) As Decimal
            If values.Count Mod 2 = 0 Then
                Return(CDec(values(values.Count \ 2 - 1)) + CDec(values(values.Count \ 2))) / 2
            Else
                Return CDec(values(values.Count \ 2))
            End If
        End Function

        ' Calculates the first and third quartiles for the specified sorted sample
        ' and returns them inside a formatted string.
        Private Function GetQuartiles(ByVal values As ArrayList) As String
            Dim part1 As ArrayList = New ArrayList()
            Dim part2 As ArrayList = New ArrayList()
            If values.Count Mod 2 = 0 Then
                part1 = values.GetRange(0, values.Count \ 2)
                part2 = values.GetRange(values.Count \ 2, values.Count \ 2)
            Else
                part1 = values.GetRange(0, values.Count \ 2 + 1)
                part2 = values.GetRange(values.Count \ 2, values.Count \ 2 + 1)
            End If

            Return String.Format("({0}, {1})", GetMedian(part1), GetMedian(part2))
        End Function
    End Class
End Namespace
