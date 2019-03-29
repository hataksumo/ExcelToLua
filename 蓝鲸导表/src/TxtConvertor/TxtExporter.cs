using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    static class TxtExporter
    {
        public static OptData getExportContent(ExcelToMapData v_data, int v_optCode,
            string rootPath, string fileName)
        {
            OptData rtn = new OptData();
            StringBuilder sb;
            try
            {
                sb = GetTxtTable(v_data);
            }
            catch (Exception ex)
            {
                rtn.errList.Add("导出基础数据时出现错误，错误信息为:\r\n" + ex.ToString());
                return rtn;
            }

            rtn.content = sb.ToString();
            string opt_path = rootPath + fileName;
            File.WriteAllText(opt_path, rtn.content);

            return rtn;
        }

        private static string[] Separator = {"#", "|",","};
        private static int GetCellTxt(ExcelMapData data, StringBuilder v_sb, int v_layer)
        {
            if (v_layer > Separator.Length)
            {
                Debug.Error("层数超过了上限");
            }
            if (data.IsLeafe)
            {
                v_sb.Append(data.LeafVal.GetTxtValue());
                return 0;
            }
            List<KeyValue<ExcelMapData>> childDatas = data.GetKeyValues();
            if(childDatas == null)
            {
                return 0;
            }
            int deep = 0;
            for (int i = 0; i < childDatas.Count; i++)  
            {
                if (i > 0)
                    v_sb.Append(Separator[deep-1]);
                deep = Math.Max(deep,GetCellTxt(childDatas[i].val, v_sb, v_layer + 1) + 1);
            }
            return deep;
        }

        private static int _findChild(string v_name, List<KeyValue<ExcelMapData>> v_childDatas)
        {
            for (int i = 0; i < v_childDatas.Count; i++)
                if (v_name == v_childDatas[i].key.ToString())
                    return i;
            return -1;
        }

        private static void _translate(ExcelMapData v_src, StringBuilder v_dst, TxtExportHeader[] listFieldName)
        {
            StringBuilder lineBuilder = new StringBuilder();
            StringBuilder cellBuilder = new StringBuilder();
            List<KeyValue<ExcelMapData>> childDatas = v_src.GetKeyValues();
            for (int i = 0; i < childDatas.Count; i++)      //行
            {
                lineBuilder.Clear();                
                ExcelMapData data = childDatas[i].val;
                if (data.Type != EExcelMapDataType.rowData)
                {
                    continue;
                }
                List<KeyValue<ExcelMapData>> childDatas2 = data.GetKeyValues();

                for (int j = 0; j < listFieldName.Length; j++)  //每一列
                {
                    if (lineBuilder.Length > 0)
                    {
                        lineBuilder.Append("\t");
                    }

                    int childIdx = _findChild(listFieldName[j].Name, childDatas2);

                    if (childIdx >= 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        GetCellTxt(childDatas2[childIdx].val, sb, 0);
                        lineBuilder.Append(sb); 
                    }
                    else
                    {
                        lineBuilder.Append(listFieldName[j].GetDefaultVal());
                    }
                    
                }
                v_dst.Append(lineBuilder);
                v_dst.Append("\r\n");
            }
        }

        public static StringBuilder GetTxtTable(ExcelToMapData v_data)
        {
            StringBuilder v_dst = new StringBuilder();
            List<TxtExportHeader> listFieldName = new List<TxtExportHeader>();
            //v_data.sheet_bins[0].header.GetTxtHeader(v_dst, listFieldName);



            TxtExportHeader[] headerNames =  v_data.sheet_bins[0].header.MyGetTxtHeader(listFieldName);
            string[] name = new string[headerNames.Length];
            for (int i = 0; i < headerNames.Length; i++)
                name[i] = headerNames[i].Name;
            v_dst.AppendLine(string.Join("\t", name));
            string[] headType = new string[headerNames.Length];
            for (int i = 0; i < headerNames.Length; i++)
                headType[i] = headerNames[i].Type;
            v_dst.AppendLine(string.Join("\t", headType));

            _translate(v_data._data, v_dst, headerNames);


            return v_dst;
        }
    }
}
