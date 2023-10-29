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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmparic : FuncionesVitales
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
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            List<int> lupasOff = new List<int>() { 2 };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementButton = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementCheck = desktopSession.FindElementsByClassName("TDBCheckBox");
            


            if (tipo == 1)
            {

                {
                    ElementCheck[34].Click();
                    ElementCheck[32].Click();
                    ElementCheck[37].Click();

                    foreach (var elem in ElementButton)
                    {
                        if (elem.Text == crudVariables[0])
                        {
                            elem.Click();
                        }
                        Thread.Sleep(500);
                    }
                    ElementList[3].SendKeys(crudVariables[1]);
                    ElementList[2].SendKeys(crudVariables[2]);

                    foreach (var elem in ElementButton)
                    {
                        if (elem.Text == crudVariables[3])
                        {
                            elem.Click();
                        }
                        Thread.Sleep(500);
                    }
                }

            }

            else
            {
                foreach (var elem in ElementButton)
                {
                    if (elem.Text == crudVariables[4])
                    {
                        elem.Click();
                    }
                    Thread.Sleep(500);
                }
            }
        }
    }
}
