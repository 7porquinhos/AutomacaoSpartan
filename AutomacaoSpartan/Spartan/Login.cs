using HtmlAgilityPack;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
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
                //Close();
                atualizarChromeDriver();
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
                    Close(); //Fecha o CHROME e LIMPA os CHROME_DRIVERS usados.
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
                    /*
                    string HTMLSerasa = _driver.PageSource; // Insere HTML da pagina na variavel.
                    string XpathArquivo = "";
                    if (HTMLSerasa.ToLower().Contains(Arquivo.acs_nome_arquivo.ToLower())) // Verifica se o nome do arquivo esta no HTML.
                    {
                        doc.LoadHtml(HTMLSerasa); // Carrega o HTML.

                        var htmlNodes = doc.DocumentNode.SelectNodes(Arquivo.obj_nome_arquivo).ToList(); // Cria uma lista com as TAG´s existentes na DIV.
                        foreach (var node in htmlNodes) // Vare cada item da lista.
                        {
                            foreach (var childNode in node.ChildNodes) // Vare os sub items da lista.
                            {
                                foreach (var childAnchor in childNode.ChildNodes) // Vare os sub items da lista.
                                {
                                    if (childAnchor.Name == "a" && childAnchor.InnerText.ToLower() == Arquivo.acs_nome_arquivo.ToLower()) // Valida a ancora encontrada é referente ao ARQUIVO procurado.
                                    {
                                        XpathArquivo = childAnchor.XPath; // Carrega o XPATH da TAG ANCHOR.
                                    }
                                }
                            }
                        }

                        // Clica na TAG A, para baixar o ARQUIVO.
                        element = new WebDriverWait(_driver, TimeSpan.FromSeconds(30)).Until(ExpectedConditions.ElementToBeClickable(By.XPath(XpathArquivo)));
                        _driver.FindElement(By.XPath(XpathArquivo)).Click();
                        Thread.Sleep(5000);
                    
                    }
                    */
                }

                Close(); //Fecha o CHROME e LIMPA os CHROME_DRIVERS usados.
            }
            catch (Exception ex)
            {
                Close();     
            }
            return true;
        }

        private static void atualizarChromeDriver()
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
            }
        }

        private void Close()
        {
            try
            {
                _driver.Close();
                _driver.Quit();
                _driver.Dispose();
            }
            catch (Exception ex)
            {

                foreach (var process in Process.GetProcessesByName("chromedriver"))
                {
                    process.Kill();
                }
            }

            try
            {
                foreach (var process in Process.GetProcessesByName("chromedriver"))
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
