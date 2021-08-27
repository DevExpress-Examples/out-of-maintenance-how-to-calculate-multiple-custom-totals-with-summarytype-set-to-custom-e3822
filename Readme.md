<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128578440/11.2.8%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E3822)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [MainWindow.xaml](./CS/DXPivotGrid_MultipleCustomTotals/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/DXPivotGrid_MultipleCustomTotals/MainWindow.xaml))
* [MainWindow.xaml.cs](./CS/DXPivotGrid_MultipleCustomTotals/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/DXPivotGrid_MultipleCustomTotals/MainWindow.xaml.vb))
<!-- default file list end -->
# How to calculate multiple Custom Totals with SummaryType set to Custom


<p>The following example demonstrates how to calculate and display multiple Custom Totals for a field.</p><br />
<p>In this example, two Custom Totals are implemented for the Sales Person field. The first one displays a median calculated against summary values, while the second one displays the first and third quartiles.</p><br />
<p>To accomplish this task, we create two PivotGridCustomTotal objects and set their summary type to FieldSummaryType.Custom. We also assign the Custom Totals' names to PivotGridCustomTotal.Tag properties to be able to distinguish between the Custom Totals when we calculate their values. Finally, we add the created objects to the Sales Person field's PivotGridField.CustomTotals collection and enable the Custom Totals to be displayed for this field by setting the PivotGridField.TotalsVisibility property to FieldTotalsVisibility.CustomTotals.</p><br />
<p>Custom Total values are actually calculated in the PivotGridControl.CustomCellValue event. First, the event handler prepares a list of summary values against which a Custom Total will be calculated. For this purpose, it creates a summary datasource and copies the summary values to an array. After that, the array is sorted and passed to an appropriate method that calculates a median or quartile value against the provided array. Finally, the resulting value is assigned to the event parameter's PivotCellValueEventArgs.Value property.</p><br />
<br />


<br/>


