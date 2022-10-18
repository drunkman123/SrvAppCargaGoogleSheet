using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.Logistica
{
    public class Resgate
    {
        public string idMat { get; set; }
        public string material { get; set; }
        public List<lstQtd> lstQtd { get; set; }
    }
    public class lstQtd
    {
        public string uorOpmCod { get; set; }
        public string nomeUnidade { get; set; }
        public int quantidade { get; set; }

    }
}
