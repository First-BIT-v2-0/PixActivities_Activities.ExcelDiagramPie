using System;
using BR.Core;
using BR.Core.Attributes;
using Microsoft.Office.Interop.Excel;
using Activities.ExcelDiagramPie.Properties;

namespace Activities.ExcelDiagramPie
{
    [LocalizableScreenName("ExcelDiagramPie_ScreenName", typeof(Resources))] // Имя активности, отображаемое в списке активностей и в заголовке шага
    [LocalizablePath("PathActivities", typeof(Resources))] // Путь к активности в панели "Активности"
    [LocalizableDescription("Activities_Description", typeof(Resources))] // описание активности

    [Image(typeof(ExcelPie), "Activities.ExcelDiagramPie.pie_icon.png")] //Иконка активности

    public class ExcelPie : Activity
    {

        [LocalizableScreenName("PathFile_ScreenName", typeof(Resources))]
        [LocalizableDescription("PathFile_Description", typeof(Resources))]
        [IsRequired]
        [IsFilePathChooser]
        public string pathFile { get; set; }

        private int defValue = 1;
        [LocalizableScreenName("NumberSheet_ScreenName", typeof(Resources))]
        [LocalizableDescription("NumberSheet_Description", typeof(Resources))]
        [IsRequired]

        public int numberSheet
        {
            get { return defValue; }
            set { defValue = value; }
        }

        [LocalizableScreenName("Cell1_ScreenName", typeof(Resources))]
        [LocalizableDescription("Cell1_Description", typeof(Resources))]
        [IsRequired]
        public string firstCell { get; set; }

        [LocalizableScreenName("Cell2_ScreenName", typeof(Resources))]
        [LocalizableDescription("Cell2_Description", typeof(Resources))]
        [IsRequired]

        public string secondCell { get; set; }

        public override void Execute(int? optionID)
        {
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
        }
    }
}
