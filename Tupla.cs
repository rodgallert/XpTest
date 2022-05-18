using System;
using System.Collections.Generic;
using System.Text;

namespace XpTest
{
    internal class Tupla
    {
        public string Cnpj { get; set; }
        public string TributacaoCvm { get; set; }
        public string TributacaoAnbima { get; set; }
        public string Resultado { get; set; }

        public Tupla(string cnpj, string tributacaoCvm, string tributacaoAnbima, string resultado)
        {
            Cnpj = cnpj;
            TributacaoCvm = tributacaoCvm;
            TributacaoAnbima = tributacaoAnbima;
            Resultado = resultado;
        }
    }
}
