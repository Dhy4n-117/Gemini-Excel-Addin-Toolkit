/**
 * Chart Creation Helpers
 * Create and configure Excel charts via Office JS API.
 */

/**
 * Map friendly chart type names to Excel chart types
 */
function getChartTypeMap() {
    return {
        bar: Excel.ChartType.barClustered,
        column: Excel.ChartType.columnClustered,
        line: Excel.ChartType.line,
        pie: Excel.ChartType.pie,
        scatter: Excel.ChartType.xyscatter,
        area: Excel.ChartType.area,
        doughnut: Excel.ChartType.doughnut,
        stacked_bar: Excel.ChartType.barStacked,
        stacked_column: Excel.ChartType.columnStacked,
    };
}

/**
 * Create a chart from a data range
 * @param {string} dataRange - Data range address (e.g., "A1:B10")
 * @param {string} chartType - Chart type (bar, column, line, pie, scatter, area)
 * @param {string} title - Chart title
 * @param {object} options - Additional chart options
 */
export async function createChart(dataRange, chartType = "column", title = "Chart", options = {}) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(dataRange);

        const CHART_TYPE_MAP = getChartTypeMap();
        const excelChartType = CHART_TYPE_MAP[chartType.toLowerCase()] || Excel.ChartType.columnClustered;

        const chart = sheet.charts.add(excelChartType, range, Excel.ChartSeriesBy.auto);

        chart.title.text = title;
        chart.title.format.font.size = 14;
        chart.title.format.font.bold = true;

        // Position the chart — place it to the right of the data
        chart.left = options.left || 400;
        chart.top = options.top || 20;
        chart.width = options.width || 480;
        chart.height = options.height || 300;

        // Style the chart
        if (options.showLegend !== undefined) {
            chart.legend.visible = options.showLegend;
        }

        await ctx.sync();
        return chart;
    });
}

/**
 * Delete all charts on the active sheet
 */
export async function clearCharts() {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const charts = sheet.charts;
        charts.load("count");

        await ctx.sync();

        for (let i = charts.count - 1; i >= 0; i--) {
            charts.getItemAt(i).delete();
        }

        await ctx.sync();
    });
}
