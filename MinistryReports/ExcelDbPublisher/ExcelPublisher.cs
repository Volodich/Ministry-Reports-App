using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using MinistryReports.Models.S21;

namespace MinistryReports.ExcelPublisher
{
    class ExcelPublisher 
    {
        internal ExcelWorksheet publishersWorksheet;
        internal ExcelPackage package;
        
        static string puthToFile;
        static string NameTable { get => "Publishers"; }

        public ExcelPublisher(S21Settings settings)
        {
            if(settings != null)
            {
                puthToFile = settings.PuthToExcelDbFile;
            }
            CheckConnect();
        }

        public bool CheckConnect()
        {
            FileInfo fileInfo = new FileInfo(puthToFile);
            if (fileInfo.Exists)
            {
                package = new ExcelPackage(fileInfo);
                ExcelWorksheet excelWorksheets = package.Workbook.Worksheets[NameTable];
                if (excelWorksheets != null)
                {
                    this.publishersWorksheet = excelWorksheets;
                    return true;
                }
                throw new FileLoadException($"Таблица не содержит данных возвещателей. Проверьте правильный ли Вы выбрали excel файл. Либо есть ли в нужном файле лист {NameTable}.");
            }
            throw new FileNotFoundException($"Невозможно найти файл по такому расположению: {puthToFile}");
        }


        public List<PublishersRange> GetPublishers(ExcelWorksheet worksheet)
        {
            var dataPublishers = (object[,])worksheet.Cells.Value;
            List<PublishersRange> publishers = new List<PublishersRange>(dataPublishers.GetLength(0));

            for (int i = 1; i < dataPublishers.GetLength(0); i++)
            {
                publishers.Add(new PublishersRange()
                {
                    Name = dataPublishers[i, 0]?.ToString(),
                    DateBirth = dataPublishers[i, 1]?.ToString() ?? "",
                    BuptismDate = dataPublishers[i, 2]?.ToString() ?? "",
                    Adress = dataPublishers[i, 3]?.ToString() ?? "",
                    Gender = dataPublishers[i, 4]?.ToString() ?? "",
                    Pioner = dataPublishers[i, 5]?.ToString() ?? "",
                    Mobile1 = dataPublishers[i, 6]?.ToString() ?? "",
                    Mobile2 = dataPublishers[i, 7]?.ToString() ?? "",
                    Appointment = dataPublishers[i, 9]?.ToString() ?? "",
                    Group = dataPublishers[i, 10]?.ToString() ?? "",
                    CountId = i
                });
            }
            return publishers;
        }

        public void AddPublisher(PublishersRange publisher, int startIndex)
        {
            using (var p = new ExcelPackage(package.File))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[NameTable];
                ws.InsertRow(startIndex, 1);
                ws.Cells[$"A{startIndex}"].Value = publisher.Name; // Имя Фамилия возвещателя
                ws.Cells[$"B{startIndex}"].Value = publisher.DateBirth; // Дата рождения
                ws.Cells[$"C{startIndex}"].Value = publisher.BuptismDate; // Дата крещения
                ws.Cells[$"D{startIndex}"].Value = publisher.Adress; // Адресс
                ws.Cells[$"E{startIndex}"].Value = publisher.Gender; // Пол
                ws.Cells[$"F{startIndex}"].Value = publisher.Pioner; // Пионер
                ws.Cells[$"G{startIndex}"].Value = publisher.Mobile1; // Номер телефона 1
                ws.Cells[$"H{startIndex}"].Value = publisher.Mobile2; // Номер телефона 2
                ws.Cells[$"I{startIndex}"].Value = publisher.Appointment; // Назначение (братья)

                package.Save();
            }
        }
    }
}
