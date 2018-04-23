using System.Windows;
using DXPivotGrid_MultipleCustomTotals.DataSet1TableAdapters;
using DevExpress.Xpf.PivotGrid;
using System.Collections;

namespace DXPivotGrid_MultipleCustomTotals {
    public partial class MainWindow : Window {
        SalesPersonTableAdapter salesPersonDataAdapter = new SalesPersonTableAdapter();
        public MainWindow() {
            InitializeComponent();

            // Binds the pivot grid to data.
            pivotGridControl1.DataSource = salesPersonDataAdapter.GetData();


            // Creates a PivotGridCustomTotal object that defines the Median Custom Total.
            PivotGridCustomTotal medianCustomTotal = new PivotGridCustomTotal();
            medianCustomTotal.SummaryType = FieldSummaryType.Custom;

            // Specifies a unique PivotGridCustomTotal.Tag property value 
            // that will be used to distinguish between two Custom Totals.
            medianCustomTotal.Tag = "Median";
            
            // Specifies formatting settings that will be used to display 
            // Custom Total column/row headers.
            medianCustomTotal.Format = "{0} Median";

            // Adds the Median Custom Total for the Sales Person field.
            fieldSalesPerson.CustomTotals.Add(medianCustomTotal);


            // Creates a PivotGridCustomTotal object that defines the Quartiles Custom Total.
            PivotGridCustomTotal quartileCustomTotal = new PivotGridCustomTotal();
            quartileCustomTotal.SummaryType = FieldSummaryType.Custom;

            // Specifies a unique PivotGridCustomTotal.Tag property value 
            // that will be used to distinguish between two Custom Totals.
            quartileCustomTotal.Tag = "Quartiles";

            // Specifies formatting settings that will be used to display 
            // Custom Total column/row headers.
            quartileCustomTotal.Format = "{0} Quartiles";

            // Adds the Quartiles Custom Total for the Sales Person field.
            fieldSalesPerson.CustomTotals.Add(quartileCustomTotal);


            // Enables the Custom Totals to be displayed instead of Automatic Totals.
            fieldSalesPerson.TotalsVisibility = FieldTotalsVisibility.CustomTotals;
            pivotGridControl1.RowTotalsLocation = FieldRowTotalsLocation.Far;
        }

        // Handles the CustomCellValue event. 
        // Fires for each data cell. If the processed cell is a Custom Total,
        // provides an appropriate Custom Total value.
        private void pivotGridControl1_CustomCellValue(object sender, PivotCellValueEventArgs e) {

            // Exits if the processed cell does not belong to a Custom Total.
            if (e.ColumnCustomTotal == null && e.RowCustomTotal == null) return;

            // Obtains a list of summary values against which
            // the Custom Total will be calculated.
            ArrayList summaryValues = GetSummaryValues(e);

            // Obtains the name of the Custom Total that should be calculated.
            string customTotalName = GetCustomTotalName(e);

            // Calculates the Custom Total value and assigns it to the Value event parameter.
            e.Value = GetCustomTotalValue(summaryValues, customTotalName);
        }

        // Returns the Custom Total name.
        private string GetCustomTotalName(PivotCellValueEventArgs e) {
            return e.ColumnCustomTotal != null ?
                e.ColumnCustomTotal.Tag.ToString() :
                e.RowCustomTotal.Tag.ToString();
        }

        // Returns a list of summary values against which
        // a Custom Total will be calculated.
        private ArrayList GetSummaryValues(PivotCellValueEventArgs e) {
            ArrayList values = new ArrayList();

            // Creates a summary data source.
            PivotSummaryDataSource sds = e.CreateSummaryDataSource();

            // Iterates through summary data source records
            // and copies summary values to an array.
            for (int i = 0; i < sds.RowCount; i++) {
                object value = sds.GetValue(i, e.DataField);
                if (value == null) {
                    continue;
                }
                values.Add(value);
            }

            // Sorts summary values.
            values.Sort();

            // Returns the summary values array.
            return values;
        }

        // Returns the Custom Total value by an array of summary values.
        private object GetCustomTotalValue(ArrayList values, string customTotalName) {

            // Returns a null value if the provided array is empty.
            if (values.Count == 0) {
                return null;
            }

            // If the Median Custom Total should be calculated,
            // calls the GetMedian method.
            if (customTotalName == "Median") {
                return GetMedian(values);
            }

            // If the Quartiles Custom Total should be calculated,
            // calls the GetQuartiles method.
            if (customTotalName == "Quartiles") {
                return GetQuartiles(values);
            }

            // Otherwise, returns a null value.
            return null;
        }

        // Calculates a median for the specified sorted sample.
        private decimal GetMedian(ArrayList values) {
            if ((values.Count % 2) == 0) {
                return ((decimal)(values[values.Count / 2 - 1]) +
                    (decimal)(values[values.Count / 2])) / 2;
            }
            else {
                return (decimal)values[values.Count / 2];
            }
        }

        // Calculates the first and third quartiles for the specified sorted sample
        // and returns them inside a formatted string.
        private string GetQuartiles(ArrayList values) {
            ArrayList part1 = new ArrayList();
            ArrayList part2 = new ArrayList();
            if ((values.Count % 2) == 0) {
                part1 = values.GetRange(0, values.Count / 2);
                part2 = values.GetRange(values.Count / 2, values.Count / 2);
            }
            else {
                part1 = values.GetRange(0, values.Count / 2 + 1);
                part2 = values.GetRange(values.Count / 2, values.Count / 2 + 1);
            }
            return string.Format("({0}, {1})", GetMedian(part1), GetMedian(part2));
        }
    }
}
