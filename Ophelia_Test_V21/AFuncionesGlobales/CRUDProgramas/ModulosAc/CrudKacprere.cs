using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKAcPrere : FuncionesVitales
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
        
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementText = desktopSession.FindElementsByClassName("TDBLookupComboBox");
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementPage = desktopSession.FindElementsByClassName("TPageControl");
            var ElementEdit = desktopSession.FindElementsByClassName("TDBGrid");
            var ElementButton = desktopSession.FindElementsByClassName("TDBCheckBox");

            if (tipo == 1)
            {
                ElementText[0].SendKeys(crudVariables[0]);

                foreach (var btRd in ElementButton)
                {
                    if (btRd.Text == "Exige Prerequisitos")
                    {
                        btRd.Click();
                    }
                }

            }
            if (tipo == 2)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 22, 7);
                desktopSession.Mouse.Click(null);

                desktopSession.Keyboard.SendKeys(crudVariables[1]);
                Thread.Sleep(6000);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 130, 27);
                //desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[3]);



            }
            if (tipo == 3)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 22, 7);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[4]);
                Thread.Sleep(6000);

            }
            if (tipo == 4)
            {
                rootSession = RootSession();               
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, (allFrame[0].Size.Height / 2)).DoubleClick().Click().Perform();
                
            }
            else
            {
                if ( tipo == 5)
                {
                    foreach (var btRd in ElementButton)
                    {
                        if (btRd.Text == "Exige Prerequisitos")
                        {
                            btRd.Click();
                        }
                    }
                }
                
            }


        }
    }
}
