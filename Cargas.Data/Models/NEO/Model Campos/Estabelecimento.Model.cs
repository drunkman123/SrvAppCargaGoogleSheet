using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.NEO
{
    public class Estabelecimento
    {
        public int cod { get; set; }
        public string cnpj { get; set; }
        public string inscricao_estadual { get; set; }
        public string dc_estabelecimento { get; set; }
        public string rz_estabelecimento { get; set; }
        public string endereco { get; set; }
        public string bairro { get; set; }
        public string cidade { get; set; }
        public string estado { get; set; }
        public string cep { get; set; }
        public string telefone { get; set; }
        public string fax { get; set; }
        public string contato { get; set; }
        public string cpf { get; set; }
        public string latitude { get; set; }
        public string longitude { get; set; }
    }
}
