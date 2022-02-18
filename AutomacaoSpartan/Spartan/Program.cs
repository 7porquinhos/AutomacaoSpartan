using Spartan.BLL.Util;
using Spartan.BLL.Util.Extensao;
using Spartan.Dominio.Entidades.SpartanConfig;
using Spartan.Dominio.Validacoes;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan
{
    class Program
    {

        static void Main(string[] args)
           {

            List<string> ListaDeProcessosForKill = new List<string>();
            ListaDeProcessosForKill.Add("chromedriver");
            ListaDeProcessosForKill.Add("EXCEL");

            ProcessExtension.ProcessKill(ListaDeProcessosForKill);

            UtilFiles utilFile = new UtilFiles();
            Login login = new Login();
            Perioticidade perioticidade = new Perioticidade();

            var retornoLista = utilFile.LerExcel();

            var lstFinal = perioticidade.ValidaExecucaoPerioticidade(retornoLista);

            if (login.LogarSerasa(lstFinal))
            {
                foreach (var item in lstFinal)
                {
                    utilFile.Remover(item);                
                }
                utilFile.EscreverExcel(retornoLista.Count, lstFinal);
            }

            /*
            string arquivo = ConfigurationManager.AppSettings["PathArquivo"];
            List<Process> processos = FileUtil.WhoIsLocking(arquivo);

            foreach (Process processo in processos)
            {
                processo.Kill();
            }
            */
        }
    }
}
