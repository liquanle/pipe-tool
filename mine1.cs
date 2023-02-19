using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataClient
{
    internal class mine1
    {
        string m_name;
        string m_mark;
        string m_class;

        public void setUserInfo(string name, string mark)
        {
            m_name = name;
            m_mark = mark;
        }

        public void Addclass(string clsname){
            m_class = clsname;
        }
    }
}
