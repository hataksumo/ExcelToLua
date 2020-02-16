using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    static class LuaExporter2
    {
        private static int minId = int.MaxValue;
        private static int maxId = int.MinValue;
        private static int rowCount = 0;
        private static string getFileNameNoSuffix(string fileName)
        {
            int pos = fileName.LastIndexOf('.');
            return fileName.Substring(0, pos);
        }
        private static void WriteLuaMain(ExcelToMapData v_data, int v_optCode,
            string rootPath, string fileName)
        {
            string content = File.ReadAllText(Config.templetPath + "lua_main_templet.lua");
            string fileNameNoSuffix = getFileNameNoSuffix(fileName);
            content = content.Replace("{name}", fileNameNoSuffix);
            content = content.Replace("{count}", v_data.GetFieldCount().ToString());
            content = content.Replace("{minID}", minId.ToString());
            content = content.Replace("{maxID}", maxId.ToString());

            string opt_path = rootPath + fileNameNoSuffix + ".lua";
            File.WriteAllText(opt_path, content);
        }
        private static OptData WriteLuaSubForm(ExcelToMapData v_data, int v_optCode,
            string rootPath, string fileName)
        {
            OptData rtn = new OptData();
            StringBuilder sb = new StringBuilder();
            try
            {
                LuaMap root = GetLuaTable(v_data._data);
                root.outputValue(sb, 0);
                rowCount = root.GetDataCount();
                minId = root.GetMinId();
                maxId = root.GetMaxId();
            }
            catch (Exception ex)
            {
                rtn.errList.Add("导出基础数据时出现错误，错误信息为:\r\n" + ex.ToString());
            }

            rtn.content = sb.ToString();

            string fileNameNoSuffix = getFileNameNoSuffix(fileName);
            string content = File.ReadAllText(Config.templetPath + "lua_subform_templet.lua");
            content = content.Replace("{name}", fileNameNoSuffix);
            content = content.Replace("{content}", rtn.content);

            string opt_path = rootPath + fileNameNoSuffix;
            if (false == System.IO.Directory.Exists(opt_path))
            {
                //创建pic文件夹
                System.IO.Directory.CreateDirectory(opt_path);
            }
            opt_path = opt_path + "\\" + fileNameNoSuffix + "_1.lua";
            
            File.WriteAllText(opt_path, content);

            return rtn;
        }
        public static OptData getExportContent(ExcelToMapData v_data, int v_optCode,
            string rootPath, string fileName)
        {
            OptData rtn = WriteLuaSubForm(v_data, v_optCode, rootPath, fileName);
            WriteLuaMain(v_data, v_optCode, rootPath, fileName);
            return rtn;
        }

        private static void _translate(ExcelMapData v_src, LuaTable v_dst)
        {
            List<KeyValue<ExcelMapData>> childDatas = v_src.GetKeyValues();
            for (int i = 0; i < childDatas.Count; i++)
            {
                KeyValue<ExcelMapData> child = childDatas[i];
                Key key = child.key;
                ExcelMapData data = child.val;
                switch (data.Type)
                {
                    case EExcelMapDataType.indexMap:
                        LuaMap indexMap = new LuaMap();
                        indexMap.init(true, ExportSheetBin.ROW_MAX_ELEMENT);
                        v_dst.addData(key, indexMap);
                        _translate(data, indexMap);
                        indexMap.Note = data.Note;
                        break;
                    case EExcelMapDataType.rowData:
                        LuaMap rowData = new LuaMap();
                        rowData.init(false, ExportSheetBin.ROW_MAX_ELEMENT);
                        v_dst.addData(key, rowData);
                        _translate(data, rowData);
                        rowData.Note = data.Note;
                        break;
                    case EExcelMapDataType.cellTable:
                        LuaTable cellTable;
                        if (data.IsArray)
                        {
                            cellTable = new LuaArray();
                            ((LuaArray)cellTable).init(false, true, ExportSheetBin.ROW_MAX_ELEMENT);
                        }
                        else
                        {
                            cellTable = new LuaMap();
                            ((LuaMap)cellTable).init(false, ExportSheetBin.ROW_MAX_ELEMENT);
                        }
                        v_dst.addData(key, cellTable);
                        _translate(data, cellTable);
                        cellTable.Note = data.Note;
                        break;
                    case EExcelMapDataType.cellData:
                        LuaValue leafVal = data.LeafVal.GetLuaValue();
                        v_dst.addData(key, leafVal);
                        leafVal.Note = data.Note;
                        break;
                }
            }
        }

        public static LuaMap GetLuaTable(ExcelMapData v_root)
        {
            LuaMap luaRoot = new LuaMap();
            luaRoot.init(true, ExportSheetBin.ROW_MAX_ELEMENT);
            _translate(v_root, luaRoot);
            return luaRoot;
        }

    }
}
