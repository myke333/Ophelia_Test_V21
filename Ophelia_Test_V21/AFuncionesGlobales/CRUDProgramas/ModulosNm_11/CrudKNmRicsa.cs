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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
	class CrudKNmRicsa:FuncionesVitales
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

        static public void CrudRicsa(WindowsDriver<WindowsElement> desktopSession, string tipo)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TPanel");
            var btn = desktopSession.FindElementsByClassName("TBitBtn");

			switch (tipo)
			{
                case ("0"):

                    desktopSession.Mouse.MouseMove(btn[0].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep = null;
                    Accep = PruebaCRUD.RootSession();
                    Accep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;

                case ("1"):

                    desktopSession.Mouse.MouseMove(btn[0].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> down1 = null;
                    down1 = PruebaCRUD.RootSession();
                    down1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep2 = null;
                    Accep2 = PruebaCRUD.RootSession();
                    Accep2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                    break;

                case ("2"):
                    desktopSession.Mouse.MouseMove(btn[0].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> down2 = null;
                    down2 = PruebaCRUD.RootSession();
                    down2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep3 = null;
                    Accep3 = PruebaCRUD.RootSession();
                    Accep3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    break;

                case ("3"):

                    desktopSession.Mouse.MouseMove(btn[0].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> down3 = null;
                    down3 = PruebaCRUD.RootSession();
                    down3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep4 = null;
                    Accep4 = PruebaCRUD.RootSession();
                    Accep4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    break;

                case ("4"):

                    desktopSession.Mouse.MouseMove(btn[0].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> down4 = null;
                    down4 = PruebaCRUD.RootSession();
                    down4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep5 = null;
                    Accep5 = PruebaCRUD.RootSession();
                    Accep5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    break;

                case ("5"):

                    desktopSession.Mouse.MouseMove(btn[0].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> down5 = null;
                    down5 = PruebaCRUD.RootSession();
                    down5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep6 = null;
                    Accep6 = PruebaCRUD.RootSession();
                    Accep6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    break;

            }


        }
    }
}
