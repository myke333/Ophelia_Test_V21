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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKBiIncar : FuncionesVitales
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
        /*
                public static void CrudIncarPrue(WindowsDriver<WindowsElement> desktopSession, string file)
                {

                }
        */
        static public void CrudIncar(WindowsDriver<WindowsElement> desktopSession, string proc, List<string> crudPrincipal, string file)
        {

            var Elementlist = desktopSession.FindElementsByClassName("TEdit");


            switch (proc)
            {
                case ("0"):

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 480, -30);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Lupa1 = null;
                    Lupa1 = PruebaCRUD.RootSession();
                    Lupa1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 470, 50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Lupa2 = null;
                    Lupa2 = PruebaCRUD.RootSession();
                    Lupa2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 470, 95);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Lupa3 = null;
                    Lupa3 = PruebaCRUD.RootSession();
                    Lupa3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, -50, 155);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> NumCarg = null;
                    NumCarg = PruebaCRUD.RootSession();
                    NumCarg.Keyboard.SendKeys("5");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 20, 155);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> ActAdm = null;
                    ActAdm = PruebaCRUD.RootSession();
                    ActAdm.Keyboard.SendKeys("10");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 250, 155);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Fec = null;
                    Fec = PruebaCRUD.RootSession();
                    Fec.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 200, 230);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Observ = null;
                    Observ = PruebaCRUD.RootSession();
                    Observ.Keyboard.SendKeys("ObservacionBiInSql");
                    Thread.Sleep(1000);

                    break;

                case ("1"):

                    desktopSession.Mouse.MouseMove(Elementlist[9].Coordinates, 830, 270);
                    newFunctions_4.ScreenshotNuevo("Boton Aceptar", true, file);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> AcepProc = null;
                    AcepProc = PruebaCRUD.RootSession();
                    AcepProc.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    newFunctions_4.ScreenshotNuevo("Click en Aceptar", true, file);

                    break;
            }

            /*
                        int cont = 0;

                        foreach(var item in Elementlist)
                        {
                            if(cont == 3)
                            {
                                item.SendKeys("2");
                            }
                            else
                            {
                                item.SendKeys(Convert.ToString(cont));
                            }

                            cont = cont + 1;

                            Thread.Sleep(1500);
                        }
            */

        }
    }
}
