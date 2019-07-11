using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLua
{
    class WorldsCatch
    {
        protected static WorldsCatch _instence;
        protected HashSet<string> _words;
        public static void Init()
        {
            _instence = new WorldsCatch();
        }

        public WorldsCatch()
        {
            _words = new HashSet<string>();
        }

        public bool addWords(string v_words)
        {
            return _words.Add(v_words);
        }
    }
}
