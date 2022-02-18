using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan.Dominio.Entidades.SpartanConfig
{
    public class ConfigEntity
    {
        public int id_Configuracao { get; set; }
        public int rpa_ativo { get; set; }
        public string login_rede { get; set; }
        public DateTime dta_configuracao { get; set; }
        public string acs_endereco { get; set; }
        public string acs_usuario { get; set; }
        public string acs_senha { get; set; }
        public string acs_nome_arquivo { get; set; }
        public string obj_entrar { get; set; }
        public string obj_usuario { get; set; }
        public string obj_senha { get; set; }
        public string obj_pasta { get; set; }
        public string obj_nome_arquivo { get; set; }
        public string mli_sgn_carregar { get; set; }
        public string mli_sgn_entrar { get; set; }
        public string mli_sgn_pasta { get; set; }
        public string mli_sgn_download { get; set; }
        public string mli_sgn_periodicidade { get; set; }
        public string arc_from { get; set; }
        public string arc_to { get; set; }
        public string sto_procedure { get; set; }
        public DateTime DataAtualiacao { get; set; }
    }
}
