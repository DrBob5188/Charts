# Gauge Chart
Automate creation of "gauge chart" in excel, for use in presentations/reports.

## Implemented VBA functions
### NewGaugeChart
Creates a new worksheet within the current workbook and populates it with enough default data to create a basic gauge from a combination of doughnut
and pie charts. The charts are overlayed, and selection of the doughnut chart for formatting is not easily possible. You can use the chartformat
macro
### ChartFormat
The charts are overlayed, and selection of the doughnut chart for formatting is not easily possible. You can use the chartformat macro to reflect
fill colour changes made to the cells corresponding to a particular chart series element.

## Basic Instructions
1. Run macro  NewGaugeChart
     Press ALT F8 select NewGaugechart and click run.

2. Adjust the parameters and fill colours within the GaugeValues and PointerValues  named ranges
    Change the pointer position by setting the value field contained in the pointervalues named range.
    Change the pointer width by setting the pointer field contained in the pointervalues named range. 
    Changes to numeric values are reflected immediately in the chart.
    Change the fill colours of the cells within the GaugeValues named range to change the gauge
      background face.  Colour changes require an additional step. (See step 3)

3. Run the chartformat macro to apply the changes back to the chart
    Press ALT F8 select ChartFormat and click run.  Note, you must make sure that the active sheet
    contains the chart of interest.

4. Add extra gauge segments if required.
     You can add extra segments by inserting cells (uniformly) in the first two columns of the
    GaugeValues named range. Make sure to choose the shift cells downoption.  Don't insert a row.

5. Copy the chart to you presentation.
     Select the chart and then click the Home tab of the excel ribbon menu strip.
     In the clipboard group click the dropdown next to the copy icon and select copy as Picture
     Paste into your Presentation/Document and use the picture format tools to crop out additional
      whitespace.