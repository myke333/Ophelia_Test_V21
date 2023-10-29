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
	class CrudKedGrafi:FuncionesVitales
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

        static public void CrudGrafi(WindowsDriver<WindowsElement> desktopSession, int lupas)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TEdit");
            
			switch (lupas)
			{
                case (0):
                    desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 400, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep = null;
                    Accep = PruebaCRUD.RootSession();
                    Accep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;

                case (1):
                    desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 800, 260);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep2 = null;
                    Accep2 = PruebaCRUD.RootSession();
                    Accep2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;

			}
            

            /*
            int cont = 0;

            foreach (var item in Elementlist)
            {
                if (cont == 3)
                {
                    item.SendKeys("2");
                }
                else
                {
                    item.SendKeys(Convert.ToString(cont));
                }

                cont = cont + 1;
            }
            */
        }
    }
}
