using ClosedXML.Excel;
using Spartan.BLL.Util.Arquivos;
using Spartan.Dominio.Entidades.SpartanConfig;
using Spartan.Dominio.Validacoes;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spartan.BLL.Util
{
    public class UtilFiles
    {
        public void Remover(ConfigEntity configEntity)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(configEntity.arc_to);
                foreach (FileInfo f in dir.GetFiles("*.csv"))
                {
                    if (f.Name.ToLower() == configEntity.acs_nome_arquivo.ToLower())
                    {
                        File.Delete(f.FullName);
                    }
                }
                Mover(configEntity);
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }
        }
        private void Mover(ConfigEntity configEntity)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(configEntity.arc_from);
                string destino = configEntity.arc_to;
                foreach (FileInfo f in dir.GetFiles("*.csv"))
                {
                    if (f.Name.ToLower() == configEntity.acs_nome_arquivo.ToLower())
                    {
                        File.Move(f.FullName, destino + "\\" + f.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }
        }

        public List<ConfigEntity> LerExcel()
        {
            DateTime dataOff;
            int numberOff;

            List<ConfigEntity> lstConfigEntity = new List<ConfigEntity>();
            XLWorkbook xls = new XLWorkbook();
            try
            {
                xls = new XLWorkbook(ConfigurationManager.AppSettings["PathArquivo"]);

                //var xls = new XLWorkbook(ConfigurationManager.AppSettings["PathArquivo"]);
                var planilha = xls.Worksheets.First(w => w.Name == ConfigurationManager.AppSettings["Aba"]);
                var totalLinhas = planilha.Rows().Count();

                // primeira linha é o cabecalho
                for (int l = 2; l <= totalLinhas; l++)
                {
                    lstConfigEntity.Add(new ConfigEntity
                    {
                        id_Configuracao = int.TryParse(
                            planilha.Cell(l, 1).Value.ToString(), out numberOff) ? int.Parse(planilha.Cell(l, 1).Value.ToString()) : numberOff,
                        rpa_ativo = int.TryParse(
                            planilha.Cell(l, 2).Value.ToString(), out numberOff) ? int.Parse(planilha.Cell(l, 2).Value.ToString()) : numberOff,
                        login_rede = planilha.Cell(l, 3).Value.ToString(),
                        dta_configuracao = DateTime.TryParse(planilha.Cell(l, 4).Value.ToString(), out dataOff) ? DateTime.Parse(planilha.Cell(l, 4).Value.ToString()) : dataOff,
                        acs_endereco = planilha.Cell(l, 5).Value.ToString(),
                        acs_usuario = planilha.Cell(l, 6).Value.ToString(),
                        acs_senha = planilha.Cell(l, 7).Value.ToString(),
                        acs_nome_arquivo = planilha.Cell(l, 8).Value.ToString(),
                        obj_entrar = planilha.Cell(l, 9).Value.ToString(),
                        obj_usuario = planilha.Cell(l, 10).Value.ToString(),
                        obj_senha = planilha.Cell(l, 11).Value.ToString(),
                        obj_pasta = planilha.Cell(l, 12).Value.ToString(),
                        obj_nome_arquivo = planilha.Cell(l, 13).Value.ToString(),
                        mli_sgn_carregar = planilha.Cell(l, 14).Value.ToString(),
                        mli_sgn_entrar = planilha.Cell(l, 15).Value.ToString(),
                        mli_sgn_pasta = planilha.Cell(l, 16).Value.ToString(),
                        mli_sgn_download = planilha.Cell(l, 17).Value.ToString(),
                        mli_sgn_periodicidade = planilha.Cell(l, 18).Value.ToString(),
                        arc_from = planilha.Cell(l, 19).Value.ToString(),
                        arc_to = planilha.Cell(l, 20).Value.ToString(),
                        sto_procedure = planilha.Cell(l, 21).Value.ToString()
                    });

                }
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }
            xls.Dispose();

            return lstConfigEntity;
        }

        public void EscreverExcel(int numeroLinhas, List<ConfigEntity> lstConfigEntity)
        {
            Excel.Application oExcel = new Excel.Application();

            oExcel.Visible = true;
            oExcel.WindowState = Excel.XlWindowState.xlMaximized;

            //pass that to workbook object  
            Excel.Workbook WB = oExcel.Workbooks.Open(ConfigurationManager.AppSettings["PathArquivo"], Missing.Value
                , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value
                , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            // statement get the workbookname  
            string ExcelWorkbookname = WB.Name;

            // statement get the worksheet count  
            int worksheetcount = WB.Worksheets.Count;

            Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];

            // statement get the firstworksheetname  
            string firstworksheetname = wks.Name;
            foreach (var configEntity in lstConfigEntity)
            {
                for (int i = 2; i <= numeroLinhas + 1; i++)
                {
                    var nome = wks.Cells[i, 8].Value.ToString();
                    if (wks.Cells[i, 8].Value.ToString().ToLower() == configEntity.acs_nome_arquivo.ToLower())
                    {
                        wks.Cells[i, 4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                }
            }

            //statement get the first cell value  
            /*
            var firstcellvalue = ((Excel.Range)wks.Cells[2, 4]).Value;
            wks.Cells[2, 4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            firstcellvalue = ((Excel.Range)wks.Cells[2, 4]).Value;
            */
            WB.Saved = true;
            WB.Save();

            WB.Close();
            oExcel.Quit();
        }
    }
}
