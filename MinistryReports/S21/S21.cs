using System;
using System.IO;
using System.Collections.Generic;
using MinistryReports.Models;
using MinistryReports.Models.S_21;
using MinistryReports.ViewModels;
using iTextSharp.text.pdf;

namespace MinistryReports.S_21
{
    public class S21Blank
    {
        static string PuhtTamplate { get; set; }
        static string PuthToFolderUnlaoding { get; set; }
        private string defaultPuthToFolderUnlaoding { get => System.IO.Path.Combine(System.IO.Path.GetPathRoot(Environment.CurrentDirectory), "MinistryReports", "Reports Publishers"); }

        public static void SetSettings(S21Settings settings)
        {
            PuhtTamplate = settings.PuthToTamplate;
            PuthToFolderUnlaoding = settings.PuthToFolderUnlaoding;
        }

        public static bool CheckTemplateFile()
        {
            FileInfo s21template = new FileInfo(PuhtTamplate);
            if (s21template.Exists)
                return true;
            else
                return false;
        }

        internal List<S21InfoPublisherField> GenerateInfoPublishers(object PublishersInfo)
        {
            S21InfoPublisherField s21PublisherData = new S21InfoPublisherField();
            List<S21InfoPublisherField> fieldPdfPublInfo = new List<S21InfoPublisherField>(); // Лист, подстраемый под поля pdf документа S-21. 
            // Уже содержит в нужном порядке всю информацию для заполнения документа.

            var infoPublishers = PublishersInfo as List<PublishersRange>;
            foreach (var publisher in infoPublishers)
            {
                {
                    // CheckBox Format
                    if (publisher.BuptismDate == "HB") s21PublisherData.HopeOther = "Off";
                    if (publisher.Gender == "М") s21PublisherData.MenGender = "Yes";
                    if (publisher.Gender == "Ж") s21PublisherData.WomenGender = "Yes";
                    if (publisher.Appointment == "СТАР") { s21PublisherData.AppointmentPastor = "Yes";}
                    if (publisher.Appointment == "СЛУЖ") s21PublisherData.AppointmentMinistryHelp = "Yes";
                    if (publisher.Pioner == "П") s21PublisherData.Pioner = "Yes";
                }
                s21PublisherData.Name = publisher.Name;
                s21PublisherData.DateBirthday = publisher.DateBirth;
                s21PublisherData.DateBaptism = publisher.BuptismDate;

                fieldPdfPublInfo.Add(s21PublisherData);
                s21PublisherData = new S21InfoPublisherField();
            }
            return fieldPdfPublInfo;
        }

        internal string[] PublisherInfoConvert(S21InfoPublisherField data)
        {
            string[] tempName = data.Name.Split(' ');
            string name = tempName[0] + tempName[1];
            return new string[] {
            data.Name + " ",
            data.DateBirthday,
            data.DateBaptism,
            data.MenGender,
            data.WomenGender,
            data.HopeOther,
            data.Hope144,
            data.AppointmentPastor,
            data.AppointmentMinistryHelp,
            data.Pioner};
        }

        public List<string> GenerateDataPublishers(object PublishersDataYear, string year)
        {
            var datas = PublishersDataYear as List<List<object>>;
            List<string> dataPublisher = new List<string>();

            List<object> Publications = datas[0];
            List<object> Videos = datas[1];
            List<object> Hours = datas[2];
            List<object> ReturnVisits = datas[3];
            List<object> BiblStudy = datas[4];
            List<object> Notates = datas[5];

            dataPublisher.Add(year);
            int monthCount = 12; // 12 месяцев в году.

            for (int i = 0; i < monthCount; i++)
            {
                dataPublisher.Add(Publications[i].ToString());
                dataPublisher.Add(Videos[i].ToString());
                dataPublisher.Add(Hours[i].ToString());
                dataPublisher.Add(ReturnVisits[i].ToString());
                dataPublisher.Add(BiblStudy[i].ToString());
                dataPublisher.Add(Notates[i].ToString());
            }
            return dataPublisher;
        }

        public void SetFieldPdf(object infoPubl, object dataPublYearNow, object dataPublYearLast, string puthToFolder = "default")
        {
            if (puthToFolder == "default")
            {
                puthToFolder = defaultPuthToFolderUnlaoding;
            }
            var infoPublisher = (string[])infoPubl;
            var dataPublisherNow = (List<string>)dataPublYearNow;
            var dataPublisherLast = (List<string>)dataPublYearLast;


            string serviceYearNow = dataPublisherNow[0]; // первый элемент коллекции - служебный год.
            string serviceYearLast = dataPublisherLast[0];
            int serviceYearN = 10; // порядковый номер в основной коллекции пдф документа с полями.
            int serviceYearL = 95;

            // Учим iTextSharp работать с pdf.
            BaseFont tnr = BaseFont.CreateFont(System.IO.Path.Combine(System.IO.Path.GetPathRoot(Environment.CurrentDirectory), "Ministry Reports", "Settings","Times New Roman.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED); // tnr - Times New Roman
            iTextSharp.text.Font font = new iTextSharp.text.Font(tnr, 12); // шрифт

            // Загрузили шаблон
            using (FileStream fs = new FileStream(PuhtTamplate, FileMode.Open, FileAccess.ReadWrite))
            {
                PdfReader pdfS21 = new PdfReader(fs);

                PdfStamper pdfStamper = new PdfStamper(pdfS21, new FileStream($"{puthToFolder}//{infoPublisher[0]}.pdf", FileMode.Create));
                pdfStamper.AcroFields.AddSubstitutionFont(tnr);

                var acroFields = pdfStamper.AcroFields;
                var keys = pdfS21.AcroFields.Fields.Keys;

                int i = 0; // iterator for loop
                int j = 1, k = 1;
                // i переменная для цикла внутри условной конструкции first page
                // k переменная для цикла внутри условной конструкции second page
                foreach (var key in keys)
                {
                    // Запись данных возвещателя
                    if (i < 10) // 9 пунктов в pdf документе
                    {

                        acroFields.SetField(key, infoPublisher[i]);
                    }
                    else if (i == serviceYearN)
                    {
                        acroFields.SetField(key, serviceYearNow);
                    }
                    else if (i == serviceYearL)
                    {
                        acroFields.SetField(key, serviceYearLast);
                    }
                    else if (i > 9 && i < 83) // first pages ministry
                    {
                        acroFields.SetField(key, dataPublisherNow[j]);
                        j++;
                    }
                    else if (i > 95 && i < 167) // second pages ministry 
                    {
                        acroFields.SetField(key, dataPublisherLast[k]);
                        k++;
                    }
                    i++;
                }

                pdfStamper.Close();
            }
        }
    }
}

