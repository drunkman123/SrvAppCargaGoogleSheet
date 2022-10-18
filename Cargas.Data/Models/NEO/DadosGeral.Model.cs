using System;
using System.Collections.Generic;
using System.Text;

namespace Cargas.Data.Models.NEO
{
    public class DadosGeral
    {
        public int pagina_atual { get; set; }
        public int total_paginas { get; set; }
        public List<Lista_os> lista_os { get; set; }
    }
}
