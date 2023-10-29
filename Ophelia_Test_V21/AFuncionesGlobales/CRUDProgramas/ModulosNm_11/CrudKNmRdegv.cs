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
	class CrudKNmRdegv:FuncionesVitales
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

        static public void CrudRdegv(WindowsDriver<WindowsElement> desktopSession, string tipo)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TEdit");

			switch (tipo)
			{
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[8].Coordinates, 80, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecIni = null;
                    FecIni = PruebaCRUD.RootSession();
                    FecIni.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[8].Coordinates, 180, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecFin = null;
                    FecFin = PruebaCRUD.RootSession();
                    FecFin.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[8].Coordinates, 290, 10);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecImp = null;
                    FecImp = PruebaCRUD.RootSession();
                    FecImp.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[8].Coordinates, 450, 20);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    

                    desktopSession.Mouse.MouseMove(Elementlist[6].Coordinates, 300, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Enti = null;
                    Enti = PruebaCRUD.RootSession();
                    Enti.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[6].Coordinates, 300, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Sucur = null;
                    Sucur = PruebaCRUD.RootSession();
                    Sucur.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[6].Coordinates, 350, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Cheque = null;
                    Cheque = PruebaCRUD.RootSession();
                    Cheque.Keyboard.SendKeys("Normal");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[6].Coordinates, 450, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> ValMin = null;
                    ValMin = PruebaCRUD.RootSession();
                    ValMin.Keyboard.SendKeys("100000");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[6].Coordinates, 10, 80);
                    desktopSession.Mouse.Click(null);
                    break;
			}

		}
    }
}
