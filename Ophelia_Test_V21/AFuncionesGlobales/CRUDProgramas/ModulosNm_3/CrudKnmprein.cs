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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmprein : FuncionesVitales
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

        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            if (bandera == 0)
            {
                //Ventana principal
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 140, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 48, 124);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                newFunctions_4.ScreenshotNuevo("Datos Insertados ventana info", true, file);
                Thread.Sleep(2000);
                //Ventana detalle
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 86, 59);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
              
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 48, 118);
                //desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Datos insertados a ventana detalle", true, file);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 198, 59);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Datos insertados a ventana fecha", true, file);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 59);
                desktopSession.Mouse.DoubleClick(null);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 59);
                desktopSession.Mouse.DoubleClick(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
            }

        }



        static public void CrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI1 = desktopSession.FindElementsByClassName("TDBGrid");
            //TPanel
            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 45, 28);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudDet[0]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudDet[1]);
                    break;
                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 453, 28);
                    desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudDet[2]);
                    break;
                case 2:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 449, 59);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Click a ventana detalle de contrato", true, file);
                    Thread.Sleep(1500);
                    break;
            }


        }

    }
}