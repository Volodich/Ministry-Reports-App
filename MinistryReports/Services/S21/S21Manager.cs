using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using MinistryReports.Models;
using MinistryReports.Models.S_21;
using MinistryReports.ViewModels;

namespace MinistryReports.Services
{
    public interface IS21Manager
    {
        IEnumerable<S21InfoPublisherField> GetS21InfoPublisherFields(S21Settings settings);
        string GetPublisherName(S21InfoPublisherField publisherInfo);
        string[] GetPublisherInfo(S21InfoPublisherField publisherInfo);
        IEnumerable<string> CreateMinistryDataPublisherToStringFormat(object dataYear, string year);
        void CreateDocument(object Name, IEnumerable<string> dataLast, IEnumerable<string> dataNow, string puthCreate);
    }

    public class S21Manager : IS21Manager
    {
        private readonly IS21Servise _s21Servise;

        public S21Manager(S21Settings setting)
        {
            _s21Servise = new S21Service();
        }

        public IEnumerable<S21InfoPublisherField> GetS21InfoPublisherFields(S21Settings settings)
        {
            ExcelPublisher.ExcelPublisher publisherEx = new MinistryReports.ExcelPublisher.ExcelPublisher(settings);

            publisherEx.GetPublishers(publisherEx.publishersWorksheet);

            return _s21Servise.GenerateInfoPublishers(null);
        }

        public string GetPublisherName(S21InfoPublisherField publisherInfo)
        {
            var tempName = _s21Servise.PublisherInfoToArray(publisherInfo);
            return tempName[0].Split(' ')[0] + " " + tempName[0].Split(' ')[1];
        }

        public string[] GetPublisherInfo(S21InfoPublisherField publisherInfo)
        {
            return _s21Servise.PublisherInfoToArray(publisherInfo);
        }

        public IEnumerable<string> CreateMinistryDataPublisherToStringFormat(object dataYear, string year)
        {
            return _s21Servise.GenerateDataPublishers(dataYear, year);
        }

        public void CreateDocument(object Name, IEnumerable<string> dataLast, IEnumerable<string> dataNow, string puthCreate)
        {
            _s21Servise.SetFieldPdf(Name, dataNow, dataLast);
        }

    }

}