using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelToLua
{
    static class Tools
    {
        public static Regex numReg = new Regex("^[0-9]+$");
        public static bool IsNumberic(string v_str)
        {
            return numReg.IsMatch(v_str);
        }

        public static string getColName(int v_col)
        {
            v_col = v_col + 1;
            StringBuilder sbCol = new StringBuilder();
            do
            {
                sbCol.Append((char)(v_col % 26 + 'A' - 1));
                v_col /= 26;
            } while (v_col > 0);
            char[] strArr = sbCol.ToString().ToCharArray();
            Array.Reverse(strArr);
            string rtn = new string(strArr);
            return rtn;
        }

        public static int getHex(string v_hexStr)
        {
            int sum = 0;
            int off = v_hexStr[0] == '0' && v_hexStr[1] == 'x' ? 2 : 0;
            Debug.Assert(v_hexStr.Length - off <= 4, "16进制整数最多只能4位");
            char[] chararr = v_hexStr.ToCharArray(off, v_hexStr.Length - off);
            for (int i = 0; i < chararr.Length; i++)
            {
                sum *= 16;
                if (char.IsDigit(chararr[i]))
                {
                    sum += chararr[i] - '0';
                }
                else if (char.IsUpper(chararr[i]))
                {
                    sum += chararr[i] - 'A';
                }
                else if (char.IsLower(chararr[i]))
                {
                    sum += chararr[i] - 'a';
                }
                else
                {
                    Debug.Assert(false, string.Format("{0}的第{1}个字符格式错误", v_hexStr, i + off));
                }
            }
            return sum;
        }

        public static bool canBeKey(string v_key)
        {
            if (!(v_key[0] >= 'a' && v_key[0] <= 'z' || v_key[0] >= 'A' && v_key[0] <= 'z' || v_key[0] == '_'))
                return false;
            int len = v_key.Length;
            for (int i = 0; i < len; i++)
            {
                if (!(v_key[i] >= 'a' && v_key[i] <= 'z' || v_key[i] >= 'A' && v_key[i] <= 'z' || v_key[i] == '_' || v_key[i] >= '0' && v_key[i] <= '9'))
                    return false;
            }
            return true;
        }
    }
}
