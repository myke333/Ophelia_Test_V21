using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmcarco : FuncionesVitales
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

                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementMemo = desktopSession.FindElementsByClassName("TDBMemo");
            var ElementEdit = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementGrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (tipo == 1)
            {


                ElementEdit[4].SendKeys(crudVariables[0]);
                ElementEdit[3].SendKeys(crudVariables[1]);
                ElementEdit[8].SendKeys(crudVariables[2]);
                ElementEdit[5].SendKeys(crudVariables[3]);
               
            }
            else if (tipo == 2)
            {
                desktopSession.Mouse.MouseMove(ElementGrid[0].Coordinates, 34, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);               
                desktopSession.Keyboard.SendKeys(crudVariables[5]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVariables[6]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVariables[7]);
            }
            else if (tipo == 3)
            {
                desktopSession.Mouse.MouseMove(ElementGrid[0].Coordinates, 34, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);
                desktopSession.Keyboard.SendKeys(crudVariables[8]);
            }

            else
            {
                ElementEdit[5].Clear();
                ElementEdit[5].SendKeys(crudVariables[4]);
            }
        }
    }
}
