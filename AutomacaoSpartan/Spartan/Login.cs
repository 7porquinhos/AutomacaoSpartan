using HtmlAgilityPack;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Spartan.BLL.Util.Arquivos;
using Spartan.BLL.Util.Extensao;
using Spartan.Dominio.Entidades.SpartanConfig;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Spartan
{
    public class Login
    {
        private IWebDriver _driver;
        protected string DRIVERS_PATH = AppDomain.CurrentDomain.BaseDirectory;
        HtmlDocument doc = new HtmlDocument();

        public Login() { }

        public bool LogarSerasa(List<ConfigEntity> lstConfigEntity)
        {
            try
            {
                var ArrayConfigWeb = lstConfigEntity.ToArray();
                var IniciarCrawler = ArrayConfigWeb[0];
                AtualizarChromeDriver();
                var chromeOption = new ChromeOptions();

                //chromeOption.AddArgument("--headless");
                chromeOption.AddArgument("--start-maximized");
                chromeOption.AddArgument("--no-sandbox");
                chromeOption.AddArgument("--privileged");
                chromeOption.AddArgument("--disable-setuid-sandbox");
                chromeOption.AddArgument("--disable-popup-blocking");

                // Carrega o site do SERASA.
                try
                {
                    var chromeDriverService = ChromeDriverService.CreateDefaultService(DRIVERS_PATH);
                    chromeDriverService.HideCommandPromptWindow = true;
                    _driver = new ChromeDriver(chromeDriverService, chromeOption, TimeSpan.FromMinutes(1));
                    _driver.Navigate().GoToUrl(IniciarCrawler.acs_endereco);
                    Thread.Sleep(int.Parse(IniciarCrawler.mli_sgn_carregar));
                }
                catch (Exception ex)
                {
                    RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                    RegistraLog.Log($"Erro: {ex.Message}");
                }

                // Insere USUARIO.
                var element = new WebDriverWait(_driver, TimeSpan.FromSeconds(30)).Until(ExpectedConditions.ElementToBeClickable(By.XPath(IniciarCrawler.obj_usuario)));
                _driver.FindElement(By.XPath(IniciarCrawler.obj_usuario)).SendKeys(IniciarCrawler.acs_usuario);

                // Insere SENHA.
                element = new WebDriverWait(_driver, TimeSpan.FromSeconds(30)).Until(ExpectedConditions.ElementToBeClickable(By.XPath(IniciarCrawler.obj_senha)));
                _driver.FindElement(By.XPath(IniciarCrawler.obj_senha)).SendKeys(IniciarCrawler.acs_senha);
                Thread.Sleep(2000);

                // Clica no botao LOGIN.
                _driver.FindElement(By.XPath(IniciarCrawler.obj_entrar)).Click();
                Thread.Sleep(int.Parse(IniciarCrawler.mli_sgn_entrar));

                // Clica no botao RECEBIDOS.
                element = new WebDriverWait(_driver, TimeSpan.FromSeconds(30)).Until(ExpectedConditions.ElementToBeClickable(By.XPath(IniciarCrawler.obj_pasta)));
                _driver.FindElement(By.XPath(IniciarCrawler.obj_pasta)).Click();
                Thread.Sleep(int.Parse(IniciarCrawler.mli_sgn_pasta));

                foreach (var Arquivo in lstConfigEntity)
                {
                    // Clica na DIV que busca.
                    element = new WebDriverWait(_driver, TimeSpan.FromSeconds(30)).Until(ExpectedConditions.ElementToBeClickable(By.XPath(Arquivo.obj_nome_arquivo)));
                    _driver.FindElement(By.XPath(Arquivo.obj_nome_arquivo)).Click();
                    Thread.Sleep(int.Parse(Arquivo.mli_sgn_download));
                }
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }
            Close();
            return true;
        }

        private static void AtualizarChromeDriver()
        {
            try
            {
                object path = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
                String chromeVersion = "",
                    from = "",
                    to = Directory.GetCurrentDirectory() + "\\chromedriver.exe",
                    driversPath = ConfigurationManager.AppSettings["PathChromeDrive"],
                    driverFolder = "";

                if (path != null)
                    chromeVersion = FileVersionInfo.GetVersionInfo(path.ToString()).FileVersion.Substring(0, 2);
                else
                    return;

                var dir = Directory.GetDirectories(driversPath);
                foreach (var item in dir)
                {
                    int inicialVersion = Convert.ToInt32(item.Replace(driversPath, "").Substring(2, 2));
                    int lastVersion = Convert.ToInt32(item.Replace(driversPath, "").Substring(5, 2));

                    if (inicialVersion <= Convert.ToInt32(chromeVersion) && lastVersion >= Convert.ToInt32(chromeVersion))
                    {
                        driverFolder = item;
                    }
                }
                if (driverFolder == "")
                    return;

                from = driverFolder + "\\chromedriver.exe";

                File.Copy(from, to, true);
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }
        }

        private void Close()
        {
            try
            {
                _driver.Close();
                _driver.Quit();
                _driver.Dispose();

                List<string> ListaDeProcessosForKill = new List<string>();
                ListaDeProcessosForKill.Add("chromedriver");
                ListaDeProcessosForKill.Add("EXCEL");

                ProcessExtension.ProcessKill(ListaDeProcessosForKill);
            }
            catch (Exception ex)
            {
                RegistraLog.Log(String.Format($"{"Log criado em "} : {DateTime.Now}"), "ArquivoLog");
                RegistraLog.Log($"Erro: {ex.Message}");
            }
        }
    }
}
