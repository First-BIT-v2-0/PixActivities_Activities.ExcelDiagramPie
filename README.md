# PixActivities_Activities.ExcelDiagramPie
Активность предназначена для создания круговой диаграммы по указанному диапазону столбца таблицы Excel

string numberSheet;
string pathFile;

Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
appExcel.Visible = true;

//Добавить рабочую книгу
Microsoft.Office.Interop.Excel.Workbook workBook = appExcel.Workbooks.Open(pathFile);

//Получить первый лист документа (счет начинается с 1)
Microsoft.Office.Interop.Excel.Worksheet worksheet
    = (Microsoft.Office.Interop.Excel.Worksheet)appExcel.Worksheets[numberSheet];

//Создание диаграммы
ChartObjects xlCharts = (ChartObjects)worksheet.ChartObjects(Type.Missing);
ChartObject myChart = (ChartObject)xlCharts.Add(250, 0, 450, 250);
Chart chart = myChart.Chart;
Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection
    = (Microsoft.Office.Interop.Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
Series series = seriesCollection.NewSeries();
series.Values = worksheet.get_Range(firstCell, secondCell);
chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;

//Закрыть книгу с сохранением
workBook.Save();
workBook.Close();

// Закрыть приложение
appExcel.Quit();