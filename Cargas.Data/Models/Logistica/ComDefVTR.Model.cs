using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.Logistica
{
    public class ComDefVTR
    {

        public string VTR { get; set; }
        public string id_com_def { get; set; }
        public string data_comunic { get; set; }
        public string odometro { get; set; }
        public string status { get; set; }
        public string cod_opm { get; set; }
        public string re_motorista { get; set; }
        public string obs { get; set; }
        public string situa { get; set; }
        public string motivo { get; set; }
        /*
        /*Jeito mais correto de fazer acesso e alteração nas variáveis
        private string VTR { get; set; }
        public void SetVTR(string testevtr)
        {
            VTR = testevtr;
        }
        public string GetVTR()
        {
            return VTR;
        }
        */
    }

}
