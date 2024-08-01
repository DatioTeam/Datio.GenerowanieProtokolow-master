using Soneta.Business;
using Soneta.Tools;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Datio.GenerowanieProtokolowJakoZalacznik
{
    public class ParametryEksportu : ContextBase
    {
        public ParametryEksportu(Context context) : base(context)
        {
        }

        bool eksportujFaktury;
        [Priority(1)]
        [Caption("Faktury")]
        public bool EksportujFaktury
        {
            get { return eksportujFaktury; ; }
            set { eksportujFaktury = value; }
        }

        bool eksportujProtokoly;
        [Priority(1)]
        [Caption("Protokoly")]
        public bool EksportujProtokoly
        {
            get { return eksportujProtokoly; ; }
            set { eksportujProtokoly = value; }
        }
    }
}
