using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace ExcelToLua
{
    struct ExcelPathConfig
    {
        public string path;
        public string[] sheets;
    }

    struct Output_designer_config
    {
        public string src_path;
        public string src_sheet;
        public string tar_path;
        public string tar_sheet;
        public string src_pm_key;
        public int src_head_row;
        public int src_row_beg;
        public string tar_pm_key;
        public int tar_head_row;
        public int tar_row_beg;
        public string[] src_cols;
        public string[] tar_cols;
    }


    static class Config
    {
        public static ExcelPathConfig attr_designer_Path = new ExcelPathConfig();
        public static ExcelPathConfig elo_data_path = new ExcelPathConfig();
        public static string packageName = "dTable";
        public static string cliPath = null;
        public static string servPath = null;
        public static string export_path = "";
        public static string excelPath = "";
        public static string indexPath = "";
        public static string templetPath = "";
        public static string luaCfgPath = "";
        public static string srcWordsFilePath = "Words.翻译.xlsx";
        public static bool isTestAssetPath = true;
        public static string assetPath = "";
        public static string[] copyCliPath = null;
        public static string[] copyServPath = null;
        public static bool isRealeace = false;
        public static List<Output_designer_config> designer_opt_configs;
        public static string simulator_src = ".\\战斗模拟_源数据.xlsx";
        public static string simulator_tar = ".\\战斗模拟_输出.xlsx";
        public static string[] outputFiles = null;



        public static void load()
        {
            XmlDocument xmlPathDoc = new XmlDocument();
            if (!File.Exists("config.xml"))
            {
                Debug.Error("配置文件config.xml缺失");
                return;
            }
            xmlPathDoc.Load("config.xml");
            XmlNode xmlroot = xmlPathDoc.SelectSingleNode("root");
            //读取策划数据包名
            packageName = xmlroot.SelectSingleNode("package").Attributes["name"].Value;
            //APP设置
            XmlNode appNode = xmlroot.SelectSingleNode("app");
            isRealeace = bool.Parse(appNode.Attributes["isRelease"].Value);
            //设置策划表路径
            XmlNode xmlPathNode = xmlroot.SelectSingleNode("path");
            export_path = xmlPathNode.Attributes["export"].Value;
            excelPath = xmlPathNode.Attributes["excelPath"].Value;
            indexPath = xmlPathNode.Attributes["indexPath"].Value;
            templetPath = xmlPathNode.Attributes["templetPath"].Value;
            luaCfgPath = xmlPathNode.Attributes["lua_cfg"].Value;
            assetPath = xmlPathNode.Attributes["assetPath"] != null? xmlPathNode.Attributes["assetPath"].Value:"null";

            //加载导出服务端路径和客户端路径
            string strCliPath = xmlPathNode.Attributes["cli"].Value;
            cliPath = strCliPath.Split('|')[0];
            string strSrvPath = xmlPathNode.Attributes["serv"].Value;
            servPath = strSrvPath.Split('|')[0];
            //加载拷贝路径
            XmlNode copyNode = xmlroot.SelectSingleNode("copyPath");
            string strCopyCliPath = copyNode.Attributes["cli"].Value;
            copyCliPath = strCopyCliPath.Split('|');
            string strCopySrvPath = copyNode.Attributes["serv"].Value;
            copyServPath = strCopySrvPath.Split('|');
            //修正路径
            cliPath = __rectify_folder_path(cliPath);
            servPath = __rectify_folder_path(servPath);
            for (int i = 0; i < copyCliPath.Length; i++)
            {
                copyCliPath[i] = __rectify_folder_path(copyCliPath[i]);
            }
            for (int i = 0; i < copyServPath.Length; i++)
            {
                copyServPath[i] = __rectify_folder_path(copyServPath[i]);
            }


            //检测所配置的路径是否有误
            string[] pathes = new string[2 + copyCliPath.Length + copyServPath.Length];
            pathes[0] = cliPath;
            pathes[1] = servPath;
            Array.Copy(copyCliPath, 0, pathes, 2, copyCliPath.Length);
            Array.Copy(copyServPath, 0, pathes, 2+ copyCliPath.Length, copyServPath.Length);
            for (int i = 0; i < pathes.Length; i++)
            {
                if (!Directory.Exists(pathes[i]))
                {
                    if (i < 2)
                    {
                        Directory.CreateDirectory(pathes[i]);
                    }
                    else
                    {
                        Debug.Exception("没有找到路径{0},请检查配置后重新启动软件", pathes[i]);
                        return;
                    }
                }
            }


            if (!Directory.Exists(assetPath))
            {
                Debug.Info("没有找到路径： {0},将不会对资源进行检测", assetPath);
                isTestAssetPath = false;
            }

            //设置设计表相关路径
            XmlNode attrDesignerNode = xmlroot.SelectSingleNode("attrDesigner");
            if (attrDesignerNode != null)
            {
                attr_designer_Path.path = attrDesignerNode.Attributes["path"].Value;
                attr_designer_Path.sheets = attrDesignerNode.Attributes["sheets"].Value.Split(';');
            }

            //要导出的表
            XmlNode outputFilesNode = xmlroot.SelectSingleNode("outputFiles");
            if (outputFilesNode != null)
            {
                string root = excelPath;
                if (outputFilesNode.Attributes["root"] != null)
                {
                    root = outputFilesNode.Attributes["root"].Value;
                }
                if (outputFilesNode.Attributes["srcFile"] != null)
                {
                    srcWordsFilePath = outputFilesNode.Attributes["srcFile"].Value;
                }
                XmlNodeList filesNode = outputFilesNode.ChildNodes;
                List<string> path = new List<string>();
                for (int i = 0; i < filesNode.Count; i++)
                {
                    XmlNode theFileNode = filesNode.Item(i);
                    string thePath = root + theFileNode.InnerText + ".xlsx";
                    path.Add(thePath);
                }
                outputFiles = path.ToArray();
            }

            //导出数据配置
            XmlNode xmlOptChildrenNodes = xmlroot.SelectSingleNode("designer_outputs");
            if (xmlOptChildrenNodes != null)
            {
                designer_opt_configs = new List<Output_designer_config>();
                var xmlOptChildNodes = xmlOptChildrenNodes.ChildNodes;
                for (int i = 0; i < xmlOptChildNodes.Count; i++)
                {
                    XmlElement childrenNode = (XmlElement)xmlOptChildNodes[i];
                    if (childrenNode.Name != "output") continue;
                    Output_designer_config newConfig = new Output_designer_config();
                    string[] strSrcPath = childrenNode.Attributes["src_path"].Value.Split('!');
                    newConfig.src_path = strSrcPath[0];
                    newConfig.src_sheet = strSrcPath[1];
                    string[] tarSrcPath = childrenNode.Attributes["tar_path"].Value.Split('!');
                    newConfig.tar_path = tarSrcPath[0];
                    newConfig.tar_sheet = tarSrcPath[1];
                    newConfig.src_pm_key = childrenNode.Attributes["src_pm_key"].Value;
                    newConfig.src_head_row = Convert.ToInt32(childrenNode.Attributes["src_head_row"].Value);
                    newConfig.src_row_beg = Convert.ToInt32(childrenNode.Attributes["src_row_beg"].Value);
                    newConfig.tar_pm_key = childrenNode.Attributes["tar_pm_key"].Value;
                    newConfig.tar_head_row = Convert.ToInt32(childrenNode.Attributes["tar_head_row"].Value);
                    newConfig.tar_row_beg = Convert.ToInt32(childrenNode.Attributes["tar_row_beg"].Value);
                    newConfig.src_cols = childrenNode.Attributes["src_opt_cols"].Value.Split(';');
                    newConfig.tar_cols = childrenNode.Attributes["tar_opt_cols"].Value.Split(';');
                    designer_opt_configs.Add(newConfig);
                }
            }   
            XmlNode xmlSimulator = xmlroot.SelectSingleNode("simulator");
            if (xmlSimulator != null)
            {
                if (xmlSimulator.Attributes["src"] != null)
                    simulator_src = xmlSimulator.Attributes["src"].Value;
                else
                    Debug.Exception("simulator节点没找到src属性");
                if (xmlSimulator.Attributes["opt"] != null)
                    simulator_tar = xmlSimulator.Attributes["opt"].Value;
                else
                    Debug.Exception("simulator节点没找到opt属性");
            }  
        }
        private static string __rectify_folder_path(string v_path)
        {
            if (v_path.Last<char>() != '\\' && v_path.Last<char>() != '/')
                return v_path + "\\";
            return v_path;
        }

        public static string[] CliPathes
        {
            get {
                string[] rtn = new string[1+copyCliPath.Length];
                rtn[0] = cliPath;
                Array.Copy(copyCliPath, 0, rtn, 1, copyCliPath.Length);
                return rtn;
            }
        }

        public static string[] SrvPathes
        {
            get
            {
                string[] rtn = new string[1 + copyServPath.Length];
                rtn[0] = servPath;
                Array.Copy(copyServPath, 0, rtn, 1, copyServPath.Length);
                return rtn;
            }
        }
    }
}
