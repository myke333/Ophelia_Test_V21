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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAC
{
    class CrudKacparam : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();
            if (bandera == 0)
            {
               

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 30, 87);
                desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 282, 120);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 378, 43);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 368, 64);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 368, 64);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 523, 114);
                desktopSession.Mouse.Click(null);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 368, 64);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Space);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 523, 114);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Space);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 379, 40);
                desktopSession.Mouse.DoubleClick(null);
            }
        }

        static public void Seguir(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TKactusDBgrid");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 517, 24);
            desktopSession.Mouse.DoubleClick(null);
           
        }

        static public void qbe2(WindowsDriver<WindowsElement> desktopSession)
        {


            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");

            //Coordenadas boton qbe2
            //string filterqb = "13";

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 404, 15);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");



            var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");


            Thread.Sleep(1500);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }

    }
}
