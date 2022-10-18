
using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.NEO
{
    public class Lista_os
    {
        public int id { get; set; }
        public int id_status { get; set; }
        public List<empenho_pedido_compra> empenho_pedido_compra { get; set; }
        public string dc_status { get; set; }
        public DateTime? dt_abertura { get; set; }
        public DateTime? dt_orcamento { get; set; }
        public DateTime? dt_aprovacao { get; set; }
        public DateTime? dt_entrega { get; set; }
        public DateTime? dt_finalizacao { get; set; }
        public DateTime? dt_cancelamento { get; set; }
        public DateTime? dt_rejeita { get; set; }
        public DateTime? data_atualizacao_km { get; set; }
        public int km_veiculo_os { get; set; }
        public string dc_observacao { get; set; }
        public Veiculo veiculo { get; set; }
        public Unidade unidade { get; set; }
        public Condutor condutor { get; set; }
        public Estabelecimento estabelecimento { get; set; }
        public List<Itens> itens { get; set; }
        //public string[] faturas { get; set; }
        //public string[] nfs_consumo { get; set; }
        //public string[] nfs_credenciado { get; set; }

    }
}
