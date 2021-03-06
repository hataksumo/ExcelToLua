﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EXCEL = Aspose.Cells;
using Lua = NLua;

namespace ExcelToLua
{
    abstract class CellValue
    {

        public bool IsStretch { get { return _isStretch; }set { _isStretch = value; } }
        protected bool _isStretch = false;

        protected abstract bool _OnInit(string v_strCellVal, string[] v_constraint);
        public abstract bool Equals(CellValue v_other);
        protected abstract LuaValue _OnGetLuaValue();        
        protected abstract JsonValue _OnGetJsonValue();
        protected abstract string _OnGetTxtValue();
        protected abstract string _OnGetXmlAttribute();
        public abstract string ToKeyString();
        public bool _isNull;
        public bool _isMiss;
        public bool _isEmpty;

        //protected ExcelHeaderDecorate _ehd = null;
        protected CellValue()
        {
            _isNull = false;
            _isMiss = false;
            _isEmpty = false;
        }
        public virtual LuaValue GetLuaValue()
        {
            if (_isMiss)
            {
                LuaNil rtn = new LuaNil();
                rtn.init(false);
                return rtn;
            }
            if (_isNull)
            {
                LuaNil rtn = new LuaNil();
                rtn.init(true);
                return rtn;
            }
            return _OnGetLuaValue();
        }

        public virtual JsonValue GetJsonValue()
        {
            if (_isMiss)
            {
                JsonNil rtn = new JsonNil();
                rtn.init(false);
                return rtn;
            }
            if (_isNull)
            {
                JsonNil rtn = new JsonNil();
                rtn.init(true);
                return rtn;
            }
            return _OnGetJsonValue();
        }
        public virtual string GetTxtValue()
        {
            if (_isMiss || _isNull)
            {
                return "";
            }
            return _OnGetTxtValue();
        }



        public virtual Key ToKey()
        {
            Key rtn = new Key();
            rtn.keytype = KeyType.String;
            rtn.skey = ToString();
            return rtn;
        }

        public virtual XmlAttributeVal GetXmlAttribute()
        {
            XmlAttributeVal rtn = new XmlAttributeVal();
            if (_isMiss)
            {
                rtn.isInvalid = true;
                return rtn;
            }
            if (_isNull)
            {
                rtn.isNil = true;
                return rtn;
            }
            rtn.val = _OnGetXmlAttribute();
            return rtn;
        }


        public virtual bool Init(EXCEL.Cell v_cellData,string v_default = null,string[] v_constraint = null)
        {
            string strVal;
            if (v_cellData.Value == null || string.IsNullOrEmpty(v_cellData.StringValue))
            {
                if (!string.IsNullOrEmpty(v_default))
                {
                    strVal = v_default;
                }
                else
                {
                    _isMiss = true;
                    return true;
                }
            }
            else
            {
                strVal = v_cellData.Value.ToString();
            }
            
            return Init(strVal, v_constraint);
        }

        public bool Init(string v_strVal, string[] v_constraint = null)
        {
            if (v_strVal == "[invalid]" || v_strVal == "[x]")
            {
                _isMiss = true;
                return true;
            }
            if (v_strVal == "[nil]")
            {
                _isNull = true;
                return true;
            }
            if (v_strVal == "[empty]")
            {
                _isEmpty = true;
            }
            return _OnInit(v_strVal, v_constraint);
        }


        public static CellValue CheckCellVal(ExcelHeaderDecorate v_ehd)
        {
            CellValue cv = CheckCellVal(v_ehd.Type);
            cv.IsStretch = v_ehd.IsStretch;
            //cv._ehd = v_ehd;
            return cv;
        }


        public static CellValue CheckCellVal(string v_type)
        {
            CellValue rtn = null;
            v_type = v_type.Trim();
            switch (v_type)
            {
                case "int":
                case "integer":
                    rtn = new IntVal();
                    break;
                case "string":
                    rtn = new StringVal();
                    break;
                case "cstring":
                    rtn = new CStringVal();
                    break;
                case "long":
                    rtn = new LongVal();
                    break;
                case "res":
                    rtn = new ResVal();
                    break;
                case "number":
                    rtn = new NumberVal();
                    break;
                case "prob":
                    rtn = new ProbVal();
                    break;
                case "float":
                    rtn = new FloatVal();
                    break;
                case "double":
                    rtn = new DoubleVal();
                    break;
                case "percent":
                    rtn = new PercentVal();
                    break;
                case "bool":
                    rtn = new BoolVal();
                    break;
                case "table":
                    rtn = new TableValue();
                    break;
                case "dataFromLuaFile":
                    rtn = new DataFromLuaFileValue();
                    break;
                case "intX":
                    rtn = new HexInt();
                    break;
                default:
                    rtn = new IDVal(v_type);
                    break;
            }
            return rtn;
        }
    }

    class MissVal : CellValue
    {
        public MissVal()
        {
            _isMiss = true;
        }
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            return true;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            return null;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            return null;
        }
        protected override string _OnGetTxtValue()
        {
            return "";
        }
        protected override string _OnGetXmlAttribute()
        {
            return null;
        }

        public override bool Equals(CellValue v_other)
        {
            return v_other is MissVal;
        }

        public override string ToKeyString()
        {
            return null;
        }

        public override string ToString()
        {
            return "miss";
        }

        public override Key ToKey()
        {
            Key rtn = new Key();
            rtn.keytype = KeyType.String;
            rtn.skey = "miss";
            return rtn;
        }
    }

    class IntVal : CellValue
    {
        protected int _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!int.TryParse(v_strCellVal, out _data))
            {
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            IntVal obj = v_other as IntVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaInteger rtn = new LuaInteger();
            rtn.init(_isEmpty ? -1 : _data);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonInteger rtn = new JsonInteger();
            rtn.init(_isEmpty ? -1 : _data);
            return rtn;
        }

        protected override string _OnGetTxtValue()
        {
            return ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            return _data.ToString();
        }

        public override Key ToKey()
        {
            Key rtn = new Key();
            rtn.keytype = KeyType.Integer;
            rtn.ikey = _data;
            return rtn;
        }
    }

    class HexInt : IntVal
    {
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            _data = Tools.getHex(v_strCellVal);
            return true;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaHexInteger rtn = new LuaHexInteger();
            rtn.init(_isEmpty ? -1 : _data);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonHexInteger rtn = new JsonHexInteger();
            rtn.init(_isEmpty ? -1 : _data);
            return rtn;
        }
        protected override string _OnGetTxtValue()
        {
            return "0x" + _data.ToString("X4");
        }
    }

    class StringVal : CellValue
    {
        protected string _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            _data = v_strCellVal.Replace("\\","\\\\").Replace("\r", "\\r").Replace("\"", "\\\"").Replace("\n","\\n");
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            StringVal obj = v_other as StringVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaString rtn = new LuaString();
            rtn.init(_isEmpty ? "" : _data);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonString rtn = new JsonString();
            rtn.init(_isEmpty ? "" : _data);
            return rtn;
        }
        protected override string _OnGetTxtValue()
        {
            return ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            return _data;
        }
        public override Key ToKey()
        {
            Key rtn = new Key();
            rtn.keytype = KeyType.String;
            rtn.skey = _data;
            return rtn;
        }
    }

    class ResVal : StringVal
    {
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            _data = v_strCellVal;
            bool bIsTest = Config.isTestAssetPath;
            if (_data[0] == '@')
            {
                _data = _data.Substring(1);
                bIsTest = false;
            }
            if (bIsTest)
            {
                string subPath = "";
                string prefix = "";
                if (v_constraint != null)
                {
                    subPath = v_constraint[0];
                    subPath = subPath.Last() == '\\' ? subPath : subPath + "\\";
                    if (v_constraint.Length > 1)
                    {
                        prefix = "." + v_constraint[1];
                    }
                }
                string path = System.IO.Path.GetFullPath(Config.assetPath + subPath + _data + prefix);
                if (!System.IO.File.Exists(path))
                {
                    Debug.Exception("没有找到名为 {0} 的资源", path);
                    return false;
                }
            }
            _data = _data.Replace("\\", "\\\\").Replace("\r\n", "\\r\\n").Replace("\"", "\\\"");
            return true;
        }
    }

    class LongVal : CellValue
    {
        protected long _data;

        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!long.TryParse(v_strCellVal, out _data))
            {
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            return true;
        }

        public override bool Equals(CellValue v_other)
        {
            LongVal longVal = v_other as LongVal;
            return longVal != null && longVal._data == _data;
        }

        protected override LuaValue _OnGetLuaValue()
        {
            LuaLong rtn = new LuaLong();
            rtn.init(_isEmpty ? -1 : _data);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonLong rtn = new JsonLong();
            rtn.init(_isEmpty ? -1 : _data);
            return rtn;
        }

        protected override string _OnGetTxtValue()
        {
            return ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            return _data.ToString();
        }

        public override Key ToKey()
        {
            Key rtn = new Key();
            rtn.keytype = KeyType.String;
            rtn.skey = _data.ToString();
            return rtn;
        }
    }

    class CStringVal : StringVal
    {
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!base._OnInit(v_strCellVal, v_constraint))
                return false;
            CstringMemo memo = CstringMemo.GetInstence();
            memo.AddCstring(_data);
            return true;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaCString rtn = new LuaCString();
            rtn.init(_data);
            return rtn;
        }
    }

    class IDVal : LongVal
    {
        //protected int _data;
        protected string m_type;
        public IDVal(string v_type)
        {
            m_type = v_type;
        }
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            NickNameColCatchManager nickNameColCatchManager = NickNameColCatchManager.getInstence();
            if (!nickNameColCatchManager.checkData(m_type, v_strCellVal, out _data))
            {
                Debug.Error("没有找到名为[{0}]的类型为{1}的ID", v_strCellVal, m_type);
                return false;
            }
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            IDVal obj = v_other as IDVal;
            return obj != null && obj._data == _data;
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            return _data.ToString();
        }
    }

    class EnumVal : IntVal
    {
        //protected int _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!int.TryParse(v_strCellVal,out _data))
            {
                _isNull = true;
                return false;
            }
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            EnumVal obj = v_other as EnumVal;
            return obj != null && obj._data == _data;
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            return "enum "+_data.ToString();
        }
    }

    class ListVal : CellValue
    {
        protected string m_type;
        protected bool m_successInit = false;
        protected CellValue[] _data;

        public ListVal(string v_type)
        {
            m_type = v_type;
        }

        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (_isEmpty)
                return true;
            string[] nameDatas = v_strCellVal.Split(',','#');
            string[] sdata = new string[nameDatas.Length];
            int len = sdata.Length;
            _data = new CellValue[len];
            for (int i = 0; i < len; i++)
            {
                _data[i] = CheckCellVal(m_type);
                _data[i].IsStretch = _isStretch;//传递一层
                if (!_data[i].Init(nameDatas[i],v_constraint))
                    return false;
            }
            m_successInit = true;
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            ListVal obj = v_other as ListVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            if (_isEmpty)
            {
                LuaArray rtn = new LuaArray();
                rtn.init(_isStretch, false, ExportSheetBin.ROW_MAX_ELEMENT);
                return rtn;
            }
            if (m_successInit)
            {
                LuaArray rtn = new LuaArray();
                rtn.init(_isStretch,false,ExportSheetBin.ROW_MAX_ELEMENT);
                for (int i = 0; i < _data.Length; i++)
                {
                    rtn.addData(_data[i].GetLuaValue());
                }
                return rtn;
            }
            else
            {
                return new LuaNil();
            }
        }
        protected override JsonValue _OnGetJsonValue()
        {
            if (_isEmpty)
            {
                JsonArray rtn = new JsonArray();
                rtn.init(_isStretch, ExportSheetBin.ROW_MAX_ELEMENT);
                return rtn;
            }
            if (m_successInit)
            {
                JsonArray rtn = new JsonArray();
                rtn.init(_isStretch, ExportSheetBin.ROW_MAX_ELEMENT);
                for (int i = 0; i < _data.Length; i++)
                {
                    rtn.addData(_data[i].GetJsonValue());
                }
                return rtn;
            }
            else
            {
                return new JsonNil();
            }
        }
        protected override string _OnGetTxtValue()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < _data.Length; i++)
            {
                if(sb.Length > 0)
                {
                    sb.Append("#");
                }
                sb.Append(_data[i].GetTxtValue());
            }
            return sb.ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < _data.Length; i++)
            {
                sb.Append(_data[i].ToString());
            }
            return sb.ToString();
        }
    }


    class BoolVal : CellValue
    {
        protected bool _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            byte val;
            if (!byte.TryParse(v_strCellVal, out val))
            {
                if (v_strCellVal.ToLower() == "true")
                {
                    _data = true;
                    return true;
                }
                else if (v_strCellVal.ToLower() == "false")
                {
                    _data = false;
                    return true;
                }
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            if (val == 0) _data = false;
            else _data = true;
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            BoolVal obj = v_other as BoolVal;
            return obj != null && obj._data == _data;
        }

        protected override LuaValue _OnGetLuaValue()
        {
            LuaBoolean rtn = new LuaBoolean();
            rtn.init(_data);
            return rtn;
        }

        protected override JsonValue _OnGetJsonValue()
        {
            JsonBoolean rtn = new JsonBoolean();
            rtn.init(_data);
            return rtn;
        }
        //????
        protected override string _OnGetTxtValue()
        {
            return _data.ToString().ToLower();
        }
        public override string ToKeyString()
        {
            return _data.ToString().ToLower();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }

        public override string ToString()
        {
            return _data.ToString();
        }

        public override Key ToKey()
        {
            Key rtn = new Key();
            rtn.keytype = KeyType.Integer;
            rtn.ikey = _data ? 1 : 0;
            return rtn;
        }
    }

    class ProbVal : CellValue
    {
        protected short _data10k;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!short.TryParse(v_strCellVal, out _data10k))
            {
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            if (_data10k < -1 && _data10k > 10000)
            {
                Debug.ExcelError(_data10k + "必须是万分制的数");
                return false;
            }
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            ProbVal obj = v_other as ProbVal;
            return obj != null && obj._data10k == _data10k;
        }

        protected override LuaValue _OnGetLuaValue()
        {
            LuaProb10k rtn = new LuaProb10k();
            rtn.init(_data10k);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonInteger rtn = new JsonInteger();
            rtn.init(_data10k);
            return rtn;
        }
        protected override string _OnGetTxtValue()
        {
            return _data10k.ToString();
        }
        public override string ToKeyString()
        {
            return string.Format("{0,2}%",_data10k/10000.0);
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data10k.ToString();
        }
    }

    class FloatVal : CellValue
    {
        protected float _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!float.TryParse(v_strCellVal, out _data))
            {
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            _data = (float)Math.Round(_data, 6);
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            FloatVal obj = v_other as FloatVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaFloat rtn = new LuaFloat();
            rtn.init(_isEmpty ? 0 : _data);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonFloat rtn = new JsonFloat();
            rtn.init(_isEmpty ? 0 : _data);
            return rtn;
        }
        protected override string _OnGetTxtValue()
        {
            return _data.ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }
    }


    class DoubleVal : CellValue
    {
        protected double _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!double.TryParse(v_strCellVal, out _data))
            {
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            _data = Math.Round(_data, 12);
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            DoubleVal obj = v_other as DoubleVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaDouble rtn = new LuaDouble();
            if (_isEmpty)
            {
                rtn.init(0);
            }
            else
            {
                rtn.init(_data);
            }
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonDouble rtn = new JsonDouble();
            if (_isEmpty)
            {
                rtn.init(0);
            }
            else
            {
                rtn.init(_data);
            }
            return rtn;
        }
        protected override string _OnGetTxtValue()
        {
            return _data.ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }
    }

    enum NumberType
    {
        Integer,
        Double,
    }

    class NumberVal : CellValue
    {
        protected double _data;
        protected int _iData;
        protected NumberType _type;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (int.TryParse(v_strCellVal,out _iData))
            {
                _type = NumberType.Integer;
                return true;
            }
            else if (double.TryParse(v_strCellVal, out _data))
            {
                _data = Math.Round(_data, 12);
                _type = NumberType.Double;
                return true;
            }
            Debug.ExcelError(v_strCellVal + "  数据格式不对");
            return false;
            
        }
        public override bool Equals(CellValue v_other)
        {
            NumberVal obj = v_other as NumberVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            switch (_type)
            {
                case NumberType.Integer:
                    LuaInteger iRtn = new LuaInteger();
                    iRtn.init(_isEmpty ? 0 : _iData);
                    return iRtn;
                case NumberType.Double:
                    LuaDouble rtn = new LuaDouble();
                    rtn.init(_isEmpty ? 0 : _data);
                    return rtn;
            }
            Debug.Exception("NumberVal类型错误");
            return new LuaNil();         
        }
        protected override JsonValue _OnGetJsonValue()
        {
            switch (_type)
            {
                case NumberType.Integer:
                    JsonInteger iRtn = new JsonInteger();
                    iRtn.init(_iData);
                    return iRtn;
                case NumberType.Double:
                    JsonDouble rtn = new JsonDouble();
                    rtn.init(_data);
                    return rtn;
            }
            Debug.Exception("NumberVal类型错误");
            return new JsonNil();
        }
        protected override string _OnGetTxtValue()
        {
            return _data.ToString();
        }
        public override string ToKeyString()
        {
            switch (_type)
            {
                case NumberType.Integer:
                    return _iData.ToString();
                case NumberType.Double:
                    return _data.ToString();
            }
            return "null";
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }
    }




    class PercentVal : CellValue
    {
        protected float _data;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            if (!float.TryParse(v_strCellVal, out _data))
            {
                Debug.ExcelError(v_strCellVal + "  数据格式不对");
                return false;
            }
            if (_data > 100 || _data < 0)
            {
                Debug.ExcelError(v_strCellVal + " percentVal必须是1~100的小数");
                return false;
            }
            _data = (float)Math.Round(_data / 100.0, 2);
            return true;
        }
        public override bool Equals(CellValue v_other)
        {
            PercentVal obj = v_other as PercentVal;
            return obj != null && obj._data == _data;
        }
        protected override LuaValue _OnGetLuaValue()
        {
            LuaPercent rtn = new LuaPercent();
            rtn.init(_data);
            return rtn;
        }
        protected override JsonValue _OnGetJsonValue()
        {
            JsonDouble rtn = new JsonDouble();
            rtn.init(_data);
            return rtn;
        }
        protected override string _OnGetTxtValue()
        {
            return _data.ToString();
        }
        public override string ToKeyString()
        {
            return _data.ToString();
        }

        protected override string _OnGetXmlAttribute()
        {
            return _data.ToString();
        }
    }

    class DataFromLuaFileValue : CellValue
    {
        protected string _key;
        protected LuaValue _luaval = null;
        public override bool Equals(CellValue v_other)
        {
            return _key == ((DataFromLuaFileValue)v_other)._key;
        }

        public override string ToKeyString()
        {
            return "TableFrmFile_"+ _key;
        }

        protected override LuaValue _OnGetLuaValue()
        {
            return _luaval;
        }

        protected override JsonValue _OnGetJsonValue()
        {
            return new JsonString(_luaval.ToString());

        }
        protected override string _OnGetTxtValue()
        {
            return _luaval.ToString();
        }
        protected override string _OnGetXmlAttribute()
        {
            throw new NotImplementedException();
        }

        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            _key = v_strCellVal;
            Lua.Lua lua_state = LuaState.Get_Instenct();
            object lua_data = null;
            var tmprst = lua_state.DoString("return " + v_strCellVal);
            if (tmprst == null)
            {
                Debug.ExcelError("return " + v_strCellVal + " is nil");
                return false;
            }
            lua_data = tmprst[0];
            if (lua_data == null)
            {
                _luaval = new LuaNil();
                return true;
            }
            else if (lua_data is Lua.LuaTable)
            {
                Lua.LuaTable table = lua_data as Lua.LuaTable;
                LuaMap map = new LuaMap();
                map.init(_isStretch, ExportSheetBin.ROW_MAX_ELEMENT);
                if (!fill_luatable(map, table))
                {
                    _luaval = new LuaNil();
                    return false;
                }
                _luaval = map;
                return true;
            }
            else if (lua_data is double)
            {
                _luaval = new LuaDouble(Convert.ToDouble(lua_data));
                return true;
            }
            else if (lua_data is string)
            {
                _luaval = new LuaString(Convert.ToString(lua_data));
                return true;
            }
            else if (lua_data is bool)
            {
                _luaval = new LuaBoolean(Convert.ToBoolean(lua_data));
                return true;
            }

            return false;
        }

        protected bool fill_luatable(LuaMap v_map,Lua.LuaTable v_luatable)
        {
            foreach (KeyValuePair<object, object> i in v_luatable)
            {
                LuaValue val = null;
                if (i.Value is Lua.LuaTable)
                {
                    LuaMap newTable = new LuaMap();
                    newTable.init(_isStretch, ExportSheetBin.ROW_MAX_ELEMENT);
                    fill_luatable(newTable, (Lua.LuaTable)i.Value);
                    val = newTable;
                }
                else
                {
                    if (i.Value is double)
                    {
                        val = new LuaDouble(Convert.ToDouble(i.Value));
                    }
                    else if (i.Value is string)
                    {
                        val = new LuaString(Convert.ToString(i.Value));
                    }
                    else if (i.Value is bool)
                    {
                        val = new LuaBoolean(Convert.ToBoolean(i.Value));
                    }
                    else if (i.Value == null)
                    {
                        val = new LuaNil();
                    }
                    else
                    {
                        Debug.Exception("出现了无法识别的luavalue,键是{0}", i.Key);
                    }
                }

                if (i.Key is int || i.Key is double)
                {
                    v_map.addData(Convert.ToInt32(i.Key), val);
                }
                else if (i.Key is string)
                {
                    v_map.addData(Convert.ToString(i.Key), val);
                }
            }
            return true;
        }
    }

    class TableValue : CellValue
    {
        protected string _source;
        protected override bool _OnInit(string v_strCellVal, string[] v_constraint)
        {
            _source = "{" + v_strCellVal + "}";
            //Lua.Lua lua_state = LuaState.Get_Instenct();
            //var lua_data = lua_state.DoString(_source)[0];
            return true;
        }

        public override bool Equals(CellValue v_other)
        {
            return _source == ((TableValue)v_other)._source;
        }

        public override String ToKeyString()
        {
            return "TableValue: " + _source;
        }

        protected override LuaValue _OnGetLuaValue()
        {
            LuaMap mp = new LuaMap();
            mp.init(false, ExportSheetBin.ROW_MAX_ELEMENT);
            return mp;
        }

        protected override JsonValue _OnGetJsonValue()
        {
            JsonMap mp = new JsonMap();
            mp.init(false, ExportSheetBin.ROW_MAX_ELEMENT);
            return mp;
        }
        protected override string _OnGetTxtValue()
        {
            throw new NotImplementedException("未实现");
        }
        protected override string _OnGetXmlAttribute()
        {
            throw new NotImplementedException();
        }
    }

}
