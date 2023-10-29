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
	class CrudKNmReapd:FuncionesVitales
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

        static public void CrudReapd(WindowsDriver<WindowsElement> desktopSession, string tipo)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TEdit");

			switch (tipo)
			{
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, -200, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Resol = null;
                    Resol = PruebaCRUD.RootSession();
                    Resol.Keyboard.SendKeys("100");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, -30, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecResol = null;
                    FecResol = PruebaCRUD.RootSession();
                    FecResol.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 100, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecIni = null;
                    FecIni = PruebaCRUD.RootSession();
                    FecIni.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 140, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> DiasTom = null; ;
                    DiasTom = PruebaCRUD.RootSession();
                    DiasTom.Keyboard.SendKeys("5");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, -200, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Observ = null;
                    Observ = PruebaCRUD.RootSession();
                    Observ.Keyboard.SendKeys("Prueba");
                    break;

                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 700, 120);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 600, 325);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Accep1 = null;
                    Accep1 = PruebaCRUD.RootSession();
                    Accep1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Accep2 = null;
                    Accep2 = PruebaCRUD.RootSession();
                    Accep2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[0].Coordinates, 800, 325);
                    desktopSession.Mouse.Click(null);
                    break;
			}
		}
    }
}
