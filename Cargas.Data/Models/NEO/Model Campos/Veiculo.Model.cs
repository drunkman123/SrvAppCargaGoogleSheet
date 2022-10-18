using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.NEO
{
    public class Veiculo
    {

        public int id { get; set; }
        public string prefixo { get; set; }
        public string placa { get; set; }
        public String km { get; set; }
        public int id_marca { get; set; }
        public string dc_marca { get; set; }
        public int id_modelo { get; set; }
        public string dc_modelo { get; set; }
        public string ano_fabricacao { get; set; }


    }
}
