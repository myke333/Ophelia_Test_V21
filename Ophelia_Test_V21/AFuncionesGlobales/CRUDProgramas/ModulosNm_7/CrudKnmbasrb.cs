using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7
{
    class CrudKnmBasrb : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex

                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
            var Elementlist1 = desktopSession.FindElementsByName(data[3]);
            var Elementlist2 = desktopSession.FindElementsByName(data[4]);
            switch (tipo)
            {
                case ("0"):
                    Elementlist[10].SendKeys(data[0]);
                    Elementlist[10].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[9].SendKeys(data[1]);
                    Elementlist[9].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[5].SendKeys(data[2]);
                    break;
                case ("1"):
                    Elementlist1[0].Click();
                    break;
                case ("2"):
                    Elementlist2[0].Click();
                    break;
            }
        }

        public static List<String> preliminar(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string nomPrograma)
        {

            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;

            for (int i = 0; i < 3; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, "1", file, pdf1);
                rootSession = RootSession();
                var ventana = rootSession.FindElementByClassName("TFrmNmImpre");

                ventana.Click();
                if (i == 0)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                }
                else if (i == 1)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }
                else
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }

                //if (i!=2)
                //{
                //    FuncionesGlobales.ReportePreliminar(desktopSession, "1", file, pdf1);
                //}

            }

            return null;
        }
    }
}