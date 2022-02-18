using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan.BLL.Util.Extensao
{
    public static class CoversorDataHora
    {
        public static TimeSpan DifencaHoras(this DateTime data1, DateTime data2)
        {
            return data2.Date - data1.Date;
        }
    }
}
