using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Interactions;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;

namespace APITest
{
    public class APISelenium
    {
        ChromeDriver driver;
        APIFuncionesVitales FV = new APIFuncionesVitales();
        String mainWin;
        String auxModalWin;

        public APISelenium()
        {
        }

        public void LoginApps(string App, string User, string Password, string UrlApp, string file) //LOGIN SELFSERVICE Y RL
        {
            var options = new ChromeOptions();
            options.AddArgument("-no-sandbox");
            driver = new ChromeDriver(@"C:\deployment\", options, TimeSpan.FromSeconds(240));
            SetImplicitTimeoutSeconds(240);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(240);
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(UrlApp);

            Thread.Sleep(500);
            FV.Screenshot("Abrir programa", true, file);
            if (App.ToLower() == "selfservice" || App.ToLower() == "smartpeople")
            {
                SendKeys("//input[contains(@name,'txtCodUsua')]", User);
                Click("//input[contains(@name,'txtPasUsua')]");                
                SendKeys("//input[contains(@name,'txtPasUsua')]", Keys.Enter);
                SendKeys("//input[contains(@name,'txtPasUsua')]", Password);
                SendKeys("//input[contains(@name,'txtPasUsua')]", Password);
                Keyboard.SendKeys("{TAB}");
            }
            else if (App.ToLower() == "reclutamiento")
            {
                //Clic a ventana emergente
                if (ExistControl("//input[contains(@id,'BtnOk')]"))
                {
                    FV.Screenshot("Política de Cookies", true, file);
                    Click("//input[contains(@id,'BtnOk')]");
                }
                Click("//label[contains(@id,'lblLogin')]");
                SendKeys("//input[contains(@id,'txtMail')]", User);
                //SendKeys("//input[contains(@id,'txtMail')]", Keys.Enter);
                Thread.Sleep(500);
                Click("//input[contains(@id,'txtPassword')]");
                Thread.Sleep(500);
                SendKeys("//input[contains(@id,'txtPassword')]", Password);
                //SendKeys("//input[contains(@id,'txtPassword')]", Password);
            }
            Thread.Sleep(500);
            FV.Screenshot("Login", true, file);

            driver.FindElement(By.XPath("//input[contains(@id,'btnIngresar')]")).Click();
        }

        public void Click(string control) //DA CLICK SOBRE EL CONTROL ENVIADO
        {
            driver.FindElement(By.XPath(control)).Click();
        }

        public void SendKeys(string control, string value) //ESCRIBE EN UNA CAJA DE TEXTO
        {
            //driver.FindElement(By.XPath(control)).Click();
            driver.FindElement(By.XPath(control)).Clear();
            driver.FindElement(By.XPath(control)).SendKeys(value);
        }

        public void Scroll(string control) //HACE SCROLL HASTA QUE ENCUENTE EL CONTROL
        {
            IWebElement endScroll = driver.FindElement(By.XPath(control));
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", endScroll);
        }

        public void ScrollTo(string x, string y)
        {
            var js = String.Format($"window.scrollTo({x}, {y})");
            IJavaScriptExecutor scroll = driver as IJavaScriptExecutor;
            scroll.ExecuteScript(js);
        }

        public void SelectElementByName(string control, string value) //SELECCIONA EL NOMBRE DE UNA LISTA DESPLEGABLE
        {
            SelectElement select = new SelectElement(driver.FindElement(By.XPath(control)));
            select.SelectByText(value);
        }

        public void ChangeMainWindow() //CAMBIA A LA VENTANA PRINCIPAL
        {
            driver.SwitchTo().Window(mainWin);
        }

        public void ChangeAuxWindow() //CAMBIA A UNA VENTANA AUXILIAR SI LA HAY
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            mainWin = windowHandles[0];
            auxModalWin = windowHandles[windowHandles.Count - 1];
        }

        public void ChangeWindow(String window)
        {
            driver.SwitchTo().Window(window);
        }

        public void MaximizeWindow()
        {
            driver.Manage().Window.Maximize();
        }

        public int CountWindow()
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            return windowHandles.Count();
        }

        public String MainWindow()
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            return windowHandles[0];
        }

        public String PopupWindow()
        {
            ReadOnlyCollection<String> windowHandles = driver.WindowHandles;
            return windowHandles[windowHandles.Count - 1];
        }

        public bool IsPresent(string control) //DEVUELVE BOOLEANO SI ESTA PRESENTE EL CONTROL A BUSCAR
        {
            return driver.FindElement(By.XPath(control)).Displayed;
        }

        public string GetText(string control) //OBTIENE EL TEXTO SEGUN EL CONTROL INDICADO
        {
            return driver.FindElement(By.XPath(control)).Text;
        }

        public bool ExistControl(string control) //INTENTA BUSCAR EL CONTROL ESPECIFICADO
        {
            SetImplicitTimeoutSeconds(7);
            List<IWebElement> Exist = new List<IWebElement>();
            Exist.AddRange(driver.FindElements(By.XPath(control)));
            if (Exist.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public int CountControl(string control) //INTENTA BUSCAR EL CONTROL ESPECIFICADO
        {
            SetImplicitTimeoutSeconds(5);
            List<IWebElement> Exist = new List<IWebElement>();
            Exist.AddRange(driver.FindElements(By.XPath(control)));
            if (Exist.Count > 0)
            {
                return Exist.Count;
            }
            else
            {
                return 0;
            }
        }

        public void SetImplicitTimeoutSeconds(int seconds) //SETEA UN TIEMPO IMPLICITO DE ESPERA EN SEGUNDOS
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(seconds);
        }

        public void Clear(string control) //LIMPIA EL CAMPO ENVIADO COMO PARAMETRO
        {
            driver.FindElement(By.XPath(control)).Clear();
        }

        public void Close()
        {
            driver.Close();
        }

        public void Dispose()
        {
            driver.Dispose();
        }

        public void Quit()
        {
            driver.Quit();
        }

        public void AcceptAlert()
        {
            driver.SwitchTo().Alert().Accept();
        }

        public ChromeDriver returnDriver()
        {
            return driver;
        }

        public string GetTextFromTextBox(string control)
        {
            return driver.FindElement(By.XPath(control)).GetAttribute("value");
        }

        public void Enter(string control)
        {
            driver.FindElement(By.XPath(control)).SendKeys(Keys.Enter);
        }

        public void Tab(string control)
        {
            driver.FindElement(By.XPath(control)).SendKeys(Keys.Tab);
        }

        public void ActiveElement()
        {
            driver.SwitchTo().ActiveElement();
        }

        public string Title()
        {
            string Titulo = driver.Title;
            return Titulo;
        }

        public string Subtitulo(string control)
        {
            IWebElement subtittle = driver.FindElement(By.Id(control));
            string Subtittle = subtittle.GetAttribute("data-original-title");
            return Subtittle;
        }

        public string Emergente(string control)
        {
            IWebElement subtittle = driver.FindElement(By.XPath(control));
            string Subtittle = subtittle.GetAttribute("data-original-title");
            return Subtittle;
        }

        public string EmergenteBotones(string control)
        {
            IWebElement icono = driver.FindElement(By.Id(control));
            string tittleIcono = icono.GetAttribute("title");
            return tittleIcono;
        }

        public string CamposEmergentes(string control, string Nombre, string file)
        {
            Click("//div[@id='ctl00_pBotones']/div");
            Thread.Sleep(500);
            Click(control);
            Thread.Sleep(1000);
            FV.Screenshot(Nombre, true, file);
            Thread.Sleep(100);
            string var = Emergente(control);
            Thread.Sleep(100);
            return var;
        }

        public void ValTabs(string file)
        {
            Click("//div[@id='ctl00_pBotones']/div");
            List<IWebElement> elementList = new List<IWebElement>();
            Thread.Sleep(800);
            elementList.AddRange(driver.FindElements(By.XPath("//*[contains(@name,'ctl00$ContenidoPagina$')]")));            
            if (elementList.Count > 0)
            {
                elementList[0].Click();
                Thread.Sleep(500);
                foreach (IWebElement pageEle in elementList)
                {
                    Keyboard.SendKeys("{TAB}");
                    FV.Screenshot("TAB", true, file);
                    Thread.Sleep(100);
                }
            }
        }
    }


    
}
