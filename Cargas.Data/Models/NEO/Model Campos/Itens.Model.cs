using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.NEO
{
    public class Itens
    {
        public int id_item { get; set; }
        public int id_categoria { get; set; }
        public string dc_categoria { get; set; }
        public string dc_item { get; set; }
        public int qtd { get; set; }
        public double vl_unitario { get; set; }
        public double vl_mao_de_obra { get; set; }
        public double vl_total_item { get; set; }
        public DateTime dt_garantia { get; set; }
        public string tipo { get; set; }
        public string qtd_percorrer { get; set; }
        public string garantia { get; set; }
        public string dc_marca { get; set; }
        public string id_num_serie { get; set; }
        public string dc_pos_aplicaca { get; set; }
        public int cod_os_unico { get; set; }
    }
}
