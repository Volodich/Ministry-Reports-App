using System;
using OfficeOpenXml;

namespace MinistryReports.Services.Publishers
{
    public interface IJwBook
    {
        void ConnectFile();
    }
    public abstract class JwBookExcel : IJwBook
    {
        public string Alphabet => "A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X-Y-Z";

        public abstract ExcelWorksheet Worksheet { get; set; }

        public abstract string PathToWorkBook { get;  }

        public abstract string NameTable { get; }

        public abstract void ConnectFile();

        public string[] GetColumnSymbolAsStringArray(int count)
        {
            string[] arrSymbolColumn = new string[count];
            string[] alphabetSym = Alphabet.Split('-');
            int iteratorAlphabet = 0;
            int countRepeate = -1;
            string symRepeate = String.Empty;

            for (int i = 0; i < arrSymbolColumn.Length; i++)
            {
                if (iteratorAlphabet >= alphabetSym.Length)
                {
                    iteratorAlphabet = 0;
                    countRepeate++;
                    symRepeate = alphabetSym[countRepeate];
                }
                arrSymbolColumn[i] = symRepeate + alphabetSym[iteratorAlphabet];
                iteratorAlphabet++;
            }
            return arrSymbolColumn;
        }
    }
}
