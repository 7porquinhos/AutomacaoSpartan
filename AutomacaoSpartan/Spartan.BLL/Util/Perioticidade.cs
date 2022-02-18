using Spartan.BLL.Util.Arquivos;
using Spartan.BLL.Util.Extensao;
using Spartan.Dominio.Entidades.SpartanConfig;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan.BLL.Util
{
    public class Perioticidade
    {
        public List<ConfigEntity> ValidaExecucaoPerioticidade(List<ConfigEntity> lstConfigEntities)
        {
            List<ConfigEntity> lstConfigEntityAtualizada = new List<ConfigEntity>();

            try
            {
                foreach (var item in lstConfigEntities)
                {
                    if (item.dta_configuracao.DifencaHoras(DateTime.Now).Days == 0 && !string.IsNullOrEmpty(item.mli_sgn_periodicidade))
                    {
                        lstConfigEntityAtualizada.Add(item);
                    }
                    else if (item.dta_configuracao.DifencaHoras(DateTime.Now).Days > 0 && !string.IsNullOrEmpty(item.mli_sgn_periodicidade))
                    {
                        lstConfigEntityAtualizada.Add(item);
                    }
                    else if (item.dta_configuracao.DifencaHoras(DateTime.Now).Days > 0 && string.IsNullOrEmpty(item.mli_sgn_periodicidade))
                    {
                        lstConfigEntityAtualizada.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }

            return lstConfigEntityAtualizada;
        }
    }
}
