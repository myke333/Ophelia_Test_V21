using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using OpenQA.Selenium.Chrome;
using System.Threading;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using OpenQA.Selenium;
using System.Collections.ObjectModel;
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesMTOM
{
    public class SmartPeopleMTOM : FuncionesVitales
    {
        //List<string> ErrorValidacion = new List<string>();
        string app = "SmartPeople";

        public static List<string> irUrlSmartPeople(string URL, string CodEmpleado, string PassSelf, string url2, List<string> tiposDoc, string file)
        {
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            var options = new ChromeOptions();
            options.AddArgument("-no-sandbox");
            ChromeDriver driver = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromMinutes(120));
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(120);
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(URL);

            hacerLogin(driver, CodEmpleado, PassSelf, url2, file);
            errors = validarDocumentacion(driver, tiposDoc, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            return ErrorValidacion;
        }

        public static void hacerLogin(ChromeDriver driver, string CodEmpleado, string PassSelf, string url2, string file)
        {
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Inicio SmartPeople", true, file);
            var txtUsuario = driver.FindElement(By.XPath("//input[contains(@name,'txtCodUsua')]"));
            txtUsuario.SendKeys(CodEmpleado);
            driver.FindElement(By.XPath("//input[contains(@name,'txtPasUsua')]")).SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//input[contains(@name,'txtPasUsua')]")).SendKeys(PassSelf);
            Thread.Sleep(500);
            driver.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            newFunctions_4.ScreenshotNuevo("Login", true, file);

            var options = new ChromeOptions();
            options.AddArgument("-no-sandbox");
            ChromeDriver driver2 = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromMinutes(120));
            driver2.Navigate().GoToUrl(url2);
        }

        public static List<string> validarDocumentacion(ChromeDriver driver, List<string> DocumentsLoad, string file)
        {
            List<string> ErrorValidacion = new List<string>();
            List<IWebElement> ExisteDocumento = new List<IWebElement>();
            int NumRows = 0;
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Validar si existen documentos subidos", true, file);
            ExisteDocumento.AddRange(driver.FindElements(By.XPath("//input[contains(@name,'LinkButton1')]")));
            if (ExisteDocumento.Count > 0)
            {
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Existe Documentación cargada", true, file);
                driver.FindElement(By.XPath("//input[contains(@name,'LinkButton1')]")).Click();
            }
            else
            {
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("No se encontro Documentación cargada", true, file);
                ErrorValidacion.Add("No se encontro educacion formal");//buscar con una variable el titulo de la pagina y
                //ponerlo asi
                //String CampoVariable= driver.FindElement(By.Xpath("//span[contains(@id,'Titulo')]"));
                // ErrorValidacion.Add("No se encontro {CampoVariable}");
            }
            string OpenDocument;
            string[] split;
            string DownloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";

            IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver;
            js1.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");

            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Documentos adjuntados desde Ophelia", true, file);

            ////EXISTEN DOCUMENTOS //////
            List<IWebElement> ExistenDocumentos = new List<IWebElement>();
            ExistenDocumentos.AddRange(driver.FindElements(By.XPath("//a[contains(text(),'Ver')]")));

            if (ExistenDocumentos.Count > 0)
            {
                for (int i = 1; i <= ExisteDocumento.Count; i++)
                {
                    ////VISUALIZA CADA DOCUMENTO Y LO ELIMINA DE DESCARGAS//////
                    js1.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");

                    driver.FindElement(By.XPath("(//a[contains(text(),'Ver')])[" + i + "]")).Click();

                    ReadOnlyCollection<String> windowHandles1 = driver.WindowHandles;
                    String mainWin1 = windowHandles1[0];
                    String modalWin1 = windowHandles1[windowHandles1.Count - 1];
                    if (windowHandles1.Count == 2)
                    {
                        driver.SwitchTo().Window(modalWin1);
                        split = DocumentsLoad[i - 1].Split('\\');
                        OpenDocument = DownloadPath + "Archivo" + split[split.Length - 1];
                        Thread.Sleep(2000);
                        Process.Start(OpenDocument);
                        Thread.Sleep(6000);
                        newFunctions_4.ScreenshotNuevo("Archivo Descargado SelfService", true, file);
                        LimpiarProcesos();
                        Thread.Sleep(1000);
                        File.Delete(OpenDocument);
                        driver.Close();
                    }
                    driver.SwitchTo().Window(mainWin1);
                }
                ExisteDocumento.Clear();
                //Proceso Guardar
                //llenar campo ciudad (DivPoli)
                //darle click al chulito de guardar

                // Volver a darle click al elemento
                //driver.FindElement(By.XPath("//input[contains(@name,'LinkButton1')]")).Click();
                //Vuelve a scrolear

                ////ELIMINA UN DOCUMENTO DESDE EL SELFSERVICE//////
                driver.FindElement(By.XPath("//a[contains(text(),'Eliminar')]")).Click();
                NumRows = 2;
                //ValidaDescarga.Clear();
                //ValidaDescarga = ValidaAdjuntos(CodEmpleado, Database, User, Table, CodEmpresa, DocumentsLoad, Almacenamiento, RutaCP, RutaFinalFTP, NumRows);
                //ValidaDescarga.ForEach(u => ErrorValidacion.Add(u));
                js1.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Eliminar documento", true, file);
                ////QUEDAN DOS DOCUMENTOS//////
                ExisteDocumento.AddRange(driver.FindElements(By.XPath("//a[contains(text(),'Ver')]")));
                if (ExisteDocumento.Count != 2)
                {
                    ErrorValidacion.Add("El numero de documentos existentes despues de eliminar en el Self Service deben ser 2 y se encontraron: " + ExisteDocumento.Count);
                }
            }
            else
            {
                ErrorValidacion.Add("No se encontraron los registros realizados desde Ophelia");
            }
            return ErrorValidacion;
        }
    }
}
