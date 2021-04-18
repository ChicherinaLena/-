using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace Parser
{
    class ExcelClass
    {
        public static IEnumerable<Metric> EnumerateMetrics(string xlsxpath)
        {

            using (var workbook = new XLWorkbook(xlsxpath))
            {

                var worksheet = workbook.Worksheets.Worksheet(1);
                int length;
                var rows = worksheet.RangeUsed().RowsUsed().Skip(2).ToArray();
                if (rows.Length - (Metric.rowscount * 50 + 3) < 50)
                {
                    length = rows.Length + 3;

                    Metric.maxrowscount = false;
                }
                else
                {
                    length = (Metric.rowscount + 1) * 50 + 3;
                }
                for (int i = Metric.rowscount * 50 + 3; i < length; i++)
                {
                    var metric = new Metric
                    {
                        Id = worksheet.Cell(i, 1).GetValue<int>(),
                        Name = worksheet.Cell(i, 2).GetValue<string>(),
                        Description = worksheet.Cell(i, 3).GetValue<string>(),
                        Source = worksheet.Cell(i, 4).GetValue<string>(),
                        Object = worksheet.Cell(i, 5).GetValue<string>(),
                        Confidentiality = worksheet.Cell(i, 6).GetValue<bool>() ? "да" : "нет",
                        Integrity = worksheet.Cell(i, 7).GetValue<bool>() ? "да" : "нет",
                        Availability = worksheet.Cell(i, 8).GetValue<bool>() ? "да" : "нет",
                        PublicationDate = worksheet.Cell(i, 9).GetValue<DateTime>(),
                        ChangesDate = worksheet.Cell(i, 10).GetValue<DateTime>(),

                    };
                    yield return metric;
                }
            }

        }
        public static string Find(string xlsxpath, string choice)
        {
            using (var workbook = new XLWorkbook(xlsxpath))
            {

                var worksheet = workbook.Worksheets.Worksheet(1);

                var rows = worksheet.RangeUsed().RowsUsed().Skip(2);

                foreach (var row in rows)
                {
                    if (choice == Convert.ToString(row.Cell(1).Value))
                    {
                        var metric = new Metric
                        {
                            Id = Convert.ToInt32(row.Cell(1).Value),
                            Name = Convert.ToString(row.Cell(2).Value),
                            Description = Convert.ToString(row.Cell(3).Value),
                            Source = Convert.ToString(row.Cell(4).Value),
                            Object = Convert.ToString(row.Cell(5).Value),
                            Confidentiality = Convert.ToBoolean(row.Cell(6).Value) ? "да" : "нет",
                            Integrity = Convert.ToBoolean(row.Cell(7).Value) ? "да" : "нет",
                            Availability = Convert.ToBoolean(row.Cell(8).Value) ? "да" : "нет",
                            PublicationDate = Convert.ToDateTime(row.Cell(9).Value),
                            ChangesDate = Convert.ToDateTime(row.Cell(10).Value),

                        };

                        return "Идентификатор угрозы " + metric.Id + "\n\n\rНаименование угрозы\n\n\r" + metric.Name + "\n\n\rОписание угрозы\n\n\r" + metric.Description + "\n\n\rИсточник угрозы\n\n\r" + metric.Source + "\n\n\rОбъект воздействия угрозы\n\n\r" + metric.Object + "\n\n\n\rНарушение конфиденциальности\n\n\r" + metric.Confidentiality + "\n\n\rНарушение целостности\n\n\r" + metric.Integrity + "\n\n\rНарушение доступности\n\n\r" + metric.Availability + "\n\n\rВремя добавления\n\n\r" + metric.PublicationDate.ToString() + "\n\n\rВремя изменения\n\n\r" + metric.ChangesDate.ToString();
                    }

                }
                return "Error";
            }
        }
        public static IEnumerable<ShortMetric> EnumerateMetricsShort(string xlsxpath)
        {

            using (var workbook = new XLWorkbook(xlsxpath))
            {
                var worksheet = workbook.Worksheets.Worksheet(1);

                var rows = worksheet.RangeUsed().RowsUsed().Skip(2);

                foreach (var row in rows)
                {
                    var metric = new ShortMetric
                    {
                        Id = "УБИ." + Convert.ToString(row.Cell(1).Value),
                        Name = Convert.ToString(row.Cell(2).Value),
                    };


                    yield return metric;
                }
            }

        }
    }

    public class Metric
        {

            public static int rowscount = 0;
            public static bool maxrowscount = true;


            public int Id { get; set; }          //        a.Идентификатор угрозы;
            public string Name { get; set; }        //        b.Наименование угрозы;
            public string Description { get; set; }               //        c.Описание угрозы;
            public string Source { get; set; }        //        d.Источник угрозы;
            public string Object { get; set; }        //        e.Объект воздействия угрозы;
            public string Confidentiality { get; set; }        //f.Нарушение конфиденциальности(да\нет);
            public string Integrity { get; set; }       //        g.Нарушение целостности(да\нет);
            public string Availability { get; set; }         //        h.Нарушение доступности(да\нет).
            public DateTime PublicationDate { get; set; }
            public DateTime ChangesDate { get; set; }

            public override string ToString()
            {
                return "Идентификатор угрозы" + Id + "/n/rНаименование угрозы" + Name + "/n/rОписание угрозы" + Description + "/n/rИсточник угрозы" + Source + "/n/rОбъект воздействия угрозы" + Object + "/n/rНарушение конфиденциальности" + Confidentiality + "/n/rНарушение целостности" + Integrity + "/n/rНарушение доступности" + Availability + "/n/rВремя добавления" + PublicationDate + "/n/rВремя изменения" + ChangesDate;
            }
        }
        public class ShortMetric
        {
            public string Id { get; set; }          //        a.Идентификатор угрозы;
            public string Name { get; set; }
        }
    }

