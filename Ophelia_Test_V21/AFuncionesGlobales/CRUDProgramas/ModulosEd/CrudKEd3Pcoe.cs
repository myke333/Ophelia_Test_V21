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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKEd3Pcoe : FuncionesVitales
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

        public static void CrudPcoe(WindowsDriver<WindowsElement> desktopSession, int tipo)
		{
            var BuscadorBtn = desktopSession.FindElementsByClassName("TBitBtn");

			switch (tipo)
			{
                case (0):
                    desktopSession.Mouse.MouseMove(BuscadorBtn[1].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    break;

                case (1):
                    WindowsDriver<WindowsElement> Enter = null;
                    Enter = PruebaCRUD.RootSession();
                    Enter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(10000);
                    Enter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(10000);
                    Enter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(10000);
                    Enter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(10000);
                    Enter.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                    break;
			}
		}
		

    }

}
