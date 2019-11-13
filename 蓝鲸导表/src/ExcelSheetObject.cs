using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Aspose.Cells;

namespace ExcelToLua
{
    public class ExcelHeaderHelp
    {
        protected Dictionary<string, int> m_header;
        public ExcelHeaderHelp()
        {
            m_header = new Dictionary<string, int>();
        }
        public void init(Excel.Worksheet v_sheet, int v_header_row = 0)
        {
            Excel.Cells data = v_sheet.Cells;
            for (int i = 0; i < 100; i++)
            {
                string header = Convert.ToString(data[v_header_row, i].Value);
                if (string.IsNullOrEmpty(header))
                    break;
                m_header.Add(header, i);
            }
        }
        public int get_col(string v_header_name)
        {
            if (m_header.ContainsKey(v_header_name))
                return m_header[v_header_name];
            //ZFDebug.Error(string.Format("没找到名为{}的列", v_header_name));
            return -1;
        }
        public int Count { get { return m_header.Count; } }

        public string[] Data
        {
            get {
                return m_header.Keys.ToArray();
            }
        }

        public int this[string v_header_name]
        {
            get {
                return get_col(v_header_name);
            }
        }

    }





    class CellPoint : IEquatable<CellPoint>
    {
        public int row;
        public int col;
        public CellPoint()
        {

        }

        public CellPoint(int v_row, int v_col)
        {
            row = v_row;
            col = v_col;
        }

        public bool Equals(CellPoint v_other)
        {
            return (row == v_other.row) && (col == v_other.col);
        }

        public override int GetHashCode()
        {
            return row * 100 + col;
        }
    }

    struct ExcelSingleKeyData
    {
        public string _pmKey;
        public List<string> dataHeaders;
        public Dictionary<string, List<object>> _datas;
    }



    class ExcelSheetObject
    {
        protected Excel.Worksheet m_sheet;
        protected ExcelHeaderHelp m_header;
        protected string m_sheet_name;
        //Dictionary<CellPoint, List<object>> m_src_infos;
        Dictionary<string, List<object>> _data;
        string _pmKey;
        public ExcelSheetObject(Excel.Worksheet v_worksheet, string v_sheet_name)
        {
            m_sheet = v_worksheet;
            m_sheet_name = v_sheet_name;
            init_header();
        }
        public void init_header(int v_headrow = 0)
        {
            m_header = new ExcelHeaderHelp();
            m_header.init(m_sheet, v_headrow);
        }


        public object get_val(int v_row, int v_col)
        {
            return m_sheet.Cells[v_row, v_col].Value;
        }
        public object get_val(int v_row, string v_header_name)
        {
            int col = m_header.get_col(v_header_name);
            if (col < 0)
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的列", m_sheet_name, v_header_name));
                return null;
            }
            object val = m_sheet.Cells[v_row, col].Value;
            return m_sheet.Cells[v_row, col].Value;
        }

        public int get_vali(int v_row, string v_header_name)
        {
            int rtn = -1;
            object ortn = get_val(v_row, v_header_name);
            if (ortn == null) return rtn;
            if (ortn is double)
            {
                rtn = (int)(double)ortn;
            }
            else if (ortn is string)
            {
                if (!int.TryParse((string)ortn, out rtn))
                {
                    Debug.Error(string.Format("{0}不是int", (string)ortn));
                }
            }
            else
            {
                Debug.Error(string.Format("{0}表{1}列的数据类型无法解析", m_sheet_name, v_header_name));
            }
            return rtn;
        }

        public double get_valf(int v_row, string v_header_name)
        {
            double rtn = -1;
            object ortn = get_val(v_row, v_header_name);
            if (ortn == null) return rtn;
            if (ortn is double)
            {
                rtn = (double)ortn;
            }
            else if (ortn is string)
            {
                if (!double.TryParse((string)ortn, out rtn))
                {
                    Debug.Error(string.Format("{0}不是int", (string)ortn));
                }
            }
            else
            {
                Debug.Error(string.Format("{0}表{1}的列的数据类型无法解析", m_sheet_name, v_header_name));
            }
            return rtn;
        }



        public string get_vals(int v_row, string v_header_name)
        {
            string rtn = "";
            object ortn = get_val(v_row, v_header_name);
            if (ortn == null) return rtn;
            if (ortn is string)
            {
                rtn = (string)ortn;
            }
            else if (ortn is string)
            {
                rtn = ortn.ToString();
            }
            else
            {
                Debug.Error(string.Format("{0}{1}的列的数据类型无法解析", m_sheet_name, v_header_name));
            }
            return rtn;
        }



        public string get_string(int v_row, int v_col)
        {
            object val = get_val(v_row, v_col) as string;
            if (val == null)
                return null;
            return Convert.ToString(val);
        }
        public int get_int(int v_row, int v_col)
        {
            object val = get_val(v_row, v_col);
            if (val == null || !(val is int))
                return -1;
            return Convert.ToInt32(val);
        }
        public double get_double(int v_row, int v_col)
        {
            object val = get_val(v_row, v_col);
            if (val == null || !(val is double))
                return -1;
            return Convert.ToDouble(val);
        }
        public void set_val_by_point(int v_row, int v_col, object v_val)
        {
            m_sheet.Cells[v_row, v_col].Value = v_val;
        }
        public void set_val(string v_header_name, int v_row, object v_val)
        {
            int col = m_header.get_col(v_header_name);
            if (col < 0)
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的列", m_sheet_name, v_header_name));
                return;
            }
            m_sheet.Cells[v_row, col].Value = v_val;
        }
        public void set_vali(string v_header_name, int v_row, int v_val)
        {
            set_val(v_header_name, v_row, v_val);
        }
        public void set_valf(string v_header_name, int v_row, double v_val, int precision = 2)
        {
            set_val(v_header_name, v_row, Math.Round(v_val, precision));
        }
        public void set_vals(string v_header_name, int v_row, string v_val)
        {
            set_val(v_header_name, v_row, v_val);
        }

        public void set_valb(string v_header_name, int v_row, bool v_b)
        {
            int col = m_header.get_col(v_header_name);
            if (col < 0)
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的列", m_sheet_name, v_header_name));
                return;
            }
            m_sheet.Cells[v_row, col].Value = v_b ? true : false;
        }

        public void clear_cell(string v_header_name, int v_row)
        {
            int col = m_header.get_col(v_header_name);
            if (col < 0)
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的列", m_sheet_name, v_header_name));
                return;
            }
            m_sheet.Cells[v_row, col].Value = null;
        }
        public void clear_row(int v_row)
        {
            int cols = m_header.Count;
            for (int i = 0; i < cols; i++)
                m_sheet.Cells[v_row, i].Value = null;
        }
        Dictionary<string, int> pm_index;
        public void init_data(string v_pmkey = "id", int v_row_begin = 3)
        {
            Excel.Cells data = m_sheet.Cells;
            _pmKey = v_pmkey;
            _data = new Dictionary<string, List<object>>();
            int col = m_header.get_col(v_pmkey);
            if (col < 0)
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的列", m_sheet_name, v_pmkey));
                return;
            }
            pm_index = new Dictionary<string, int>();
            string[] headerNames = m_header.Data;
            for (int i = v_row_begin; i < 100000; i++)
            {
                object test_obj = data[i, col].Value;
                if (test_obj == null)
                {
                    break;
                }
                if (!pm_index.ContainsKey(test_obj.ToString()))
                {
                    pm_index.Add(test_obj.ToString(), i);
                    List<object> rowData = new List<object>();
                    _data[test_obj.ToString()] = rowData;
                    for (int j = 0; j < headerNames.Length; j++)
                    {
                        rowData.Add(get_val(i, m_header[headerNames[j]]));
                    }
                }

            }
        }

        public void set_val_by_pmid(string v_pmkey, string v_col_name, object v_val)
        {
            int col = m_header.get_col(v_col_name);
            if (col < 0)
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的列", m_sheet_name, v_col_name));
                return;
            }
            if (!pm_index.ContainsKey(v_pmkey))
            {
                Debug.Error(string.Format("{0}中没找到名为{1}的主键", m_sheet_name, v_pmkey));
                return;
            }
            int row = pm_index[v_pmkey];
            m_sheet.Cells[row, col].Value = v_val;
        }

        public Excel.Worksheet Sheet
        {
            get { return m_sheet; }
        }

        public ExcelHeaderHelp Header
        {
            get { return m_header; }
        }

        public ExcelSingleKeyData getExcelSingleKeyData()
        {
            ExcelSingleKeyData rtn = new ExcelSingleKeyData();
            rtn._pmKey = _pmKey;
            rtn.dataHeaders = new List<string>();
            rtn._datas = new Dictionary<string, List<object>>();
            string[] headerNames = m_header.Data;
            for (int i = 0; i < headerNames.Length; i++)
            {
                if (headerNames[i] != _pmKey)
                {
                    rtn.dataHeaders.Add(headerNames[i]);
                }
            }

            foreach (KeyValuePair<string, List<object>> pair in _data)
            {
                rtn._datas[pair.Key] = new List<object>();
                for (int i = 0; i < headerNames.Length; i++)
                {
                    if (headerNames[i] != _pmKey)
                    {
                        rtn._datas[pair.Key].Add(pair.Value[i]);
                    }
                }
            }
            return rtn;
        }
    }
}
