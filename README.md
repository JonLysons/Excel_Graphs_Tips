# Excel Graphics Tips

## Tables

### Create a table

1. Insert a table in your spreadsheet.
2. Select a cell within your data.
3. Select Home > Format as Table.
4. Choose a style for your table.
5. In the Create Table dialog box, set your cell range.
6. Mark if your table has headers.
7. Select OK.

### Sort a table

1. Select a cell within the data.
2. Select Home > Sort & Filter. Or, select Data > Sort.
3. Select an option:
   + Sort A to Z - sorts the selected column in an ascending order.
   + Sort Z to A - sorts the selected column in a descending order.
   + Custom Sort - sorts data in multiple columns by applying different sort criteria.
     – Here's how to do a custom sort:
     - 1. Select Custom Sort.
     - 2. Select Add Level.
     - 3. Add Level
     - 4. For Column, select the column you want to Sort by from the drop-down, and then select the second column you Then by want to sort. For example, Sort by Department and Then by Status.
     - 5. For Sort On, select Values.
     - 6. For Order, select an option, like A to Z, Smallest to Largest, or Largest to Smallest.
     - 7. For each additional column that you want to sort by, repeat steps 2-5. Note: To delete a level, select Delete Level.
     - 8. Check the My data has headers checkbox, if your data has a header row.
     - 9. Select OK.

### Filter a table

In Excel, you can create three kinds of filters: by values, by a format, or by criteria. But each of these filter types is mutually exclusive.
Filters are additive. This means that each additional filter is based on the current filter and further reduces the subset of data. You can make complex filters by filtering on more than one value, more than one format, or more than one criteria. 

1. Select a cell within the data.
2. On the Data tab, click Filter.
3. Click the arrow at the top of the column that contains the content that you want to filter.
4. Under Filter, click Choose One, and enter your filter criteria.
5. In the box next to the pop-up menu, enter the number you want to use. 
6. Depending on your choice, you may be offered additional criteria to select.

Notes: 
You can apply filters to only one range of cells on a sheet at a time.
When you apply a filter to a column, the only filters available for other columns are the values visible in the currently filtered range.
Only the first 10,000 unique entries in a list appear in the filter window.

Use wildcards:
1. `?` (question mark) – Any single character, eg sm?th finds "smith" and "smyth"
2. `*` (asterisk) – Any number of characters, eg *east finds "Northeast" and "Southeast"
3. `~` (tilde) – A question mark or an asterisk, eg there~? finds "there?"

### Total row

You can quickly total data in an Excel table by enabling the Total Row option, and then use one of several functions that are provided in a drop-down list for each table column. 

1. Go to Table > Total Row.
2. The Total Row is inserted at the bottom of your table.
3. Select the column you want to total, then select an option from the drop-down list.

### Slicers to filter data

Slicers provide buttons that you can click to filter tables, or PivotTables. In addition to quick filtering, slicers also indicate the current filtering state, which makes it easy to understand what exactly is currently displayed. To create a slicer:

1. Click in the Table or Pivot Table.
2. On the Insert tab, select Slicer.
3. In the Insert Slicers dialog box, select the check boxes for the fields you want to display then click OK.
4. A slicer is displayed for every field that is selected. Clicking on any of the slicer buttons will automatically apply that filter to the linked table  or Pivot Table.

To change the colour of the slicer used the Slicer tab.

To disconnect a Slicer from a Pivot Table, click the table, then click PivotTable Analyze tab and select Filter Connections. Follow the steps in the dialog box. 

To delete a slicer, select it and press Delete.

## Charts

Charts help you visualize your data in a way that creates maximum impact on your audience.

1. Select data for the chart.
2. Select Insert > Recommended Charts.
3. Select a chart on the Recommended Charts tab, to preview the chart.
4. Select a chart.
5. Select OK.

Add a trendline

1. Select a chart.
2. Select Design > Add Chart Element.
3. Select Trendline and then select the type of trendline you want, such as Linear, Exponential, Linear Forecast, or Moving Average.

When the values in a chart vary widely from data series to data series, you can plot one or more data series on a secondary axis. A secondary axis can also be used as part of a combination chart when you have mixed types of data (for example, price and volume) in the same chart.

1. On the View menu, click Print Layout.
2. In the chart, select the data series that you want to add a trendline to, and then click the Chart Design tab.
3. On the Chart Design tab, click Add Chart Element, and then click Trendline.
4. Choose a trendline option or click More Trendline Options.
5. You can choose from the following:
   + Exponential
   + Linear
   + Logarithmic
   + Polynomial
   + Power
   + Moving average

### Quick analysis charts

Quick Analysis takes a range of data and helps you pick the perfect chart with just a few commands. 
1. Select a range of cells.
2. Select the Quick Analysis button that appears at the bottom right corner of the selected data.
3. Select Charts.
4. Hover over the chart types to preview a chart, and then select the chart you want.
5. Select More > All Charts to all available see all charts available. Preview and select OK when done to insert the chart.

### Change which chart axis is emphasized

You can change the way that table rows and columns are plotted in a chart.  A chart plots the rows of data from the table on the chart's vertical (value) axis and the columns of data on the horizontal (category) axis. You can reverse the way the chart is plotted.
1. Select the chart.
2. Select Chart Design > Switch Row/Column.

### Change the order of the data series

You can change the order of a data series in a chart with more than one data series.
1. In the chart, select a data series. For example, in a column chart, click a column, and all the columns of that data series become selected.
2. Select Chart Design > Select Data.
3. In the Select Data Source dialog box, next to Legend entries (Series), use the up and down arrows to move the series up or down in the list.
4. Select OK.

### Change the fill color of the data series

1. In the chart, select a data series. For example, in a column chart, click a column, and all the columns of that data series become selected.
2. Select Format.
3. Under Chart Element Styles, select Shape Fill  Fill Color button and choose the color.

### Add data labels

You can add labels to show the data point values from the Excel sheet in the chart.
1. Select the chart, and then select Chart Design.
2. Select Add Chart Element > Data Labels.
3. Select the location for the data label (for example, select Outside End).
4. Depending on the chart type, some options may not be available.

### Add a data table

1. Select the chart., and then click the tab.
2. Select Chart Design > Add Chart Element > Data Table.
3. Select the options.

### Sparklines

A sparkline is a tiny chart in a worksheet cell that provides a visual representation of data. Use sparklines to show trends in a series of values, such as seasonal increases or decreases, economic cycles, or to highlight maximum and minimum values. Position a sparkline near its data for greatest impact.

Add a Sparkline
1. Select a blank cell at the end of a row of data.
2. Select Insert and pick Sparkline type, like Line, or Column.
3. Select cells in the row and OK in menu.
4. More rows of data? Drag handle to add a Sparkline for each row.

Format a Sparkline chart
1. Select the Sparkline chart.
2. Select Sparkline and then select an option.
   + Select Line, Column, or Win/Loss to change the chart type.
   + Check Markers to highlight individual values in the Sparkline chart.
   + Select a Style for the Sparkline.
   + Select Sparkline Color and the color.
   + Select Sparkline Color > Weight to select the width of the Sparkline.
   + Select Marker Color to change the color of the markers.
   + If the data has positive and negative values, select Axis to show the axis.

## Pivot tables
