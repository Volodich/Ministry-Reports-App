using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MinistryReports.S_21;
using MinistryReports.ViewModels;
using MinistryReports.Models.S_21;
using MinistryReports.Models;
using System.Windows.Navigation;

namespace MinistryReports.Controllers
{
    class S21Controller
    {
        public static UserSettings Settings { get; set; }
        public static S21Blank blank { get; set; }

        public S21Controller(S21Settings setting)
        {
            blank = new S21Blank();
            SetSettings(setting);
        }

        public static bool CheckTemplateFile()
        {
            return S21Blank.CheckTemplateFile();
        }

        public static void SetSettings(S21Settings settings)
        {
            S21Blank.SetSettings(settings);
        }

        public static List<S21InfoPublisherField> CreateS21ModelDataPublisher(S21Settings settings)
        {
            ExcelPublisher.ExcelPublisher publisherEx = new MinistryReports.ExcelPublisher.ExcelPublisher(settings);
            return blank.GenerateInfoPublishers(publisherEx.GetPublishers(publisherEx.publishersWorksheet));
        }

        public static Task<List<S21InfoPublisherField>> CreateS21ModelDataPublisherAsync(S21Settings settings)
        {
            return Task.Run(() => CreateS21ModelDataPublisher(settings));
        }

        public static string GetPublisherName(S21InfoPublisherField publisherInfo)
        {
            var tempName = blank.PublisherInfoConvert(publisherInfo);
            return tempName[0].Split(' ')[0] + " " + tempName[0].Split(' ')[1];
        }

        public string[] GetPublisherInfo(object publisherInfo)
        {
            var ipublisher = publisherInfo as S21InfoPublisherField;
            return blank.PublisherInfoConvert(ipublisher);
        }

        public List<string> CreateMinistryDataPublisherToStringFormat(object dataYear, string year)
        {
            return blank.GenerateDataPublishers(dataYear, year);
        }

        public static void CreateDocument(object Name, List<string> dataLast, List<string> dataNow, string puthCreate)
        {
            blank.SetFieldPdf(Name, dataNow, dataLast);
        }

        public static Task CreateDocumentAsync(object Name, List<string> dataLast, List<string> dataNow, string puthCreate)
        {
            return Task.Run(() => blank.SetFieldPdf(Name, dataNow, dataLast, puthCreate));
        }
    }
}
