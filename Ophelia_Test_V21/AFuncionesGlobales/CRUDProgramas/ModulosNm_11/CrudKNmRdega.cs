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
	class CrudKNmRdega:FuncionesVitales
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

        static public void CrudRdega(WindowsDriver<WindowsElement> desktopSession, string tipo)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TEdit");

			switch (tipo)
			{
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 80, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Des = null;
                    Des = PruebaCRUD.RootSession();
                    Des.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 180, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Has = null;
                    Has = PruebaCRUD.RootSession();
                    Has.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 280, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 490, 15);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 50, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 280, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[4].Coordinates, 490, 60);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> BanCorp = null;
                    BanCorp = PruebaCRUD.RootSession();
                    BanCorp.Keyboard.SendKeys("0");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 100, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> Lista = null;
                    Lista = PruebaCRUD.RootSession();
                    Lista.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> SelecLista = null;
                    SelecLista = PruebaCRUD.RootSession();
                    SelecLista.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, -200, -15);
                    desktopSession.Mouse.Click(null);
                    break;

                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 550, 220);
                    desktopSession.Mouse.Click(null);
                    break;
			}
        }
    }
}
