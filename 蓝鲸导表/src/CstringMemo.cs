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
        protected Dictionary<string, Dictionary<string,string>> _data;

        public CstringMemo()
        {
            _words = new HashSet<string>();
            _data = new Dictionary<string, Dictionary<string, string>>();
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
                if (_data.ContainsKey(words))
                {
                    Dictionary<string, string> language = _data[words];
                    foreach (string lanHeader in language.Keys)
                    {
                        if(lanHeader!="Id")
                            eso.set_vals(lanHeader, row, language[lanHeader]);
                    }
                }
                row = row + 1;
            }
            book.Save(v_excelPath);
        }

        public void initByFile(string v_excelPath)
        {
            Excel.Workbook book = new Excel.Workbook(v_excelPath);
            Excel.Worksheet sheet = book.Worksheets["CString"];
            ExcelSheetObject eso = new ExcelSheetObject(sheet, "CString");
            eso.init_header();
            eso.init_data("Key");
            ExcelSingleKeyData eskd = eso.getExcelSingleKeyData();
            string[] headerNames = eskd.dataHeaders.ToArray();
            foreach (string key in eskd._datas.Keys)
            {
                if (!_data.ContainsKey(key))
                {
                    _data[key] = new Dictionary<string, string>();
                    _words.Add(key);
                    for (int i = 0; i < eskd.dataHeaders.Count; i++)
                    {
                        _data[key][eskd.dataHeaders[i]] =  Convert.ToString(eskd._datas[key][i]);
                    }
                }

            }
        }

    }
}
