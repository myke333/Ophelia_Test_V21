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
    class CrudKacnivel : FuncionesVitales
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
            List<int> lupasOff = new List<int>() { 0, 1 };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementEdit = desktopSession.FindElementsByClassName("TDBGrid");
            
            var ElementPage = desktopSession.FindElementsByClassName("TPageControl");

            if (tipo == 1)
            {

                ElementList[4].SendKeys(crudVariables[0]);
                ElementList[3].SendKeys(crudVariables[1]); //editar
                ElementList[2].SendKeys(crudVariables[2]);
                ElementList[10].SendKeys(crudVariables[3]);
                ElementList[9].SendKeys(crudVariables[4]);
                ElementList[7].SendKeys(crudVariables[5]);
                ElementList[6].SendKeys(crudVariables[6]);




            }
            else if (tipo == 2)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 22, 7);
                desktopSession.Mouse.Click(null);

                desktopSession.Keyboard.SendKeys(crudVariables[7]);
                Thread.Sleep(6000);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 130, 27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[8]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVariables[9]);

                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Keyboard.SendKeys(crudVariables[9]);


                //ElementList[3].Clear();
                //ElementList[3].SendKeys(crudVariables[2]);

            }
            else if (tipo == 3)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 22, 7);
                desktopSession.Mouse.Click(null);

                desktopSession.Keyboard.SendKeys(crudVariables[10]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVariables[11]);
            }

        }
    }
}
