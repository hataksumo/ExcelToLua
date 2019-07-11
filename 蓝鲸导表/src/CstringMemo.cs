using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Aspose.Cells;

namespace ExcelToLua
{
    class CstringMemo
    {
        static CstringMemo _instence = null;
        static public CstringMemo GetInstence()
        {
            if (_instence == null)
            {
                _instence = new CstringMemo();
            }
            return _instence;
        }

        protected HashSet<string> _words;

        public CstringMemo()
        {
            _words = new HashSet<string>();
        }

        public bool AddCstring(string v_cstring)
        {
            return _words.Add(v_cstring);
        }

        public void OutputMemoExcel(string v_excelPath)
        {
            Excel.Workbook book = new Excel.Workbook(v_excelPath);
            Excel.Worksheet sheet = book.Worksheets["CString"];
            ExcelSheetObject eso = new ExcelSheetObject(sheet, "CString");
            eso.init_header();
            int row = 3;
            foreach (string words in _words)
            {
                eso.set_vali("Id", row, row - 2);
                eso.set_vals("Key", row, words);
                row = row + 1;
            }
            book.Save(v_excelPath);
        }
    }
}
