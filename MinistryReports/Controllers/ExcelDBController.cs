using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MinistryReports.Models;
using MinistryReports.ExcelPublisher;
using MinistryReports.Models.S21;
using MinistryReports.Models.JWBook;
using System.Runtime.InteropServices.WindowsRuntime;

namespace MinistryReports.Controllers
{
    class ExcelDBController
    {
        public static object GetDataPublisher(S21Settings settings)
        {
            ExcelPublisher.ExcelPublisher publishers = new ExcelPublisher.ExcelPublisher(settings);
            if (publishers.CheckConnect())
                return publishers.GetPublishers(publishers.publishersWorksheet);
            return null;
        }

        public static Task<object> GetDataPublisherAsync(S21Settings settings)
        {
            return Task.Run(() => GetDataPublisher(settings));

        }

        public static bool CheckConnect(S21Settings settings)
        {
            ExcelPublisher.ExcelPublisher publisher = new ExcelPublisher.ExcelPublisher(settings);
            return publisher.CheckConnect();
        }

        public static void AddPublisher(PublishersRange publisher, S21Settings settings)
        {
            ExcelPublisher.ExcelPublisher excel = new ExcelPublisher.ExcelPublisher(settings);
            if (excel.CheckConnect() == true)
            {
                var datas = GetDataPublisher(settings) as List<PublishersRange>;
                List<string> names = new List<string>(); // Будет хранить имена всех возвещателей, который находяться в excel.
                foreach (var data in datas)
                {
                    names.Add(data.Name);
                }
                names.Add(publisher.Name);
                names.Sort();
                int index = names.IndexOf(publisher.Name) + 2; // +2 -- есть общие начальные колонки в ексель таблице.

                excel.AddPublisher(publisher, index);
            }
        }
    }
}
