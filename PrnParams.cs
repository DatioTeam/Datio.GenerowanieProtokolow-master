using Soneta.Business;
using Soneta.CRM;
using Soneta.Tools;
using Soneta.Types;


namespace Datio.GenerowanieProtokolowJakoZalacznik
{
    public class PrnParams : ContextBase
    {
        public PrnParams(Context context) : base(context)
        {
        }

        FromTo okres = FromTo.Month(Date.Today);
        [Priority(1)]
        [Caption("Okres")]
        public FromTo Okres
        {
            get { return okres; }
            set { okres = value; }
        }

        Kontrahent kontrahent;
        [Priority(1)]
        [Caption("Kontrahent")]
        public Kontrahent Kontrahent
        {
            get { return kontrahent; }
            set { kontrahent = value; }
        }

        bool drukujTylkoZafakturowane;
        [Priority(1)]
        [Caption("Tylko zafak.")]
        public bool DrukowanieZafakturowanych
        {
            get { return drukujTylkoZafakturowane;  }
            set { drukujTylkoZafakturowane = value; }
        }
    }
}
