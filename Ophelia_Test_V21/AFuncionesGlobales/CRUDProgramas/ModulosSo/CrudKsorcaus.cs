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
using Ophelia_Test_V21.TestMethods.ModulosSO;
using OpenQA.Selenium;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKsorcaus /*: FuncionesVitales*/
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


        // no sirve la funcion predeterminada del  Qbe; se tuvo  que automatizar este paso
        static public void qbe2(WindowsDriver<WindowsElement> desktopSession)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 896, 314);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            RootSession = ReloadSession(RootSession, "TKQBE");
            var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 255, 24);
            RootSession.Mouse.DoubleClick(null);
            Thread.Sleep(1000);

            RootSession.Keyboard.SendKeys("IGE");
            Thread.Sleep(1000);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }


        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                //Fecha Inicial 
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 113, 33);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys("01/09/2022");
                Thread.Sleep(500);
 
            }
            else
            {
                ////Fecha Final 
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys("05/09/2022");
                Thread.Sleep(500);

            }

        }




        static public void ClickBtnPreli(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            desktopSession.FindElementByName("Preliminar").Click();
            Thread.Sleep(1200);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            desktopSession.Mouse.Click(null);


            RootSession = ReloadSession(RootSession, "TFrmMnSoRCAus");
            var btnTDVI2 = RootSession.FindElementsByClassName("TRadioGroup");
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 48, 19);
            RootSession.Mouse.Click(null);

            newFunctions_4.ScreenshotNuevo("Boton Preliminar1", true, file);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2300);

        }


        static public void ClickBtnPreli2(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            desktopSession.FindElementByName("Preliminar").Click();
            Thread.Sleep(1200);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            desktopSession.Mouse.Click(null);


            RootSession = ReloadSession(RootSession, "TFrmMnSoRCAus");
            var btnTDVI2 = RootSession.FindElementsByClassName("TRadioGroup");
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 48, 19);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1000);

            // Flecha hacia bajo en la ventana previa del preli.
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);

            
            // Click en la lupa para escoger opcion obligatoria.
            var btnTDVI1 = RootSession.FindElementsByClassName("TPanel");
            RootSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 234, 248);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1000);

            // Nueva.
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);

            newFunctions_4.ScreenshotNuevo("Preliminar 2", true, file);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);


        }

        static public void ClickBtnPreli3(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            desktopSession.FindElementByName("Preliminar").Click();
            Thread.Sleep(1200);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            Thread.Sleep(1200);

            RootSession = ReloadSession(RootSession, "TFrmMnSoRCAus");
            var btnTDVI2 = RootSession.FindElementsByClassName("TRadioGroup");

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 48, 19);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1000);
            // Flecha hacia bajo en la ventana previa del preli.
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            newFunctions_4.ScreenshotNuevo("Preliminar 3", true, file);
            Thread.Sleep(900);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2300);

        }

        static public void ClickBtnPreli4(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            desktopSession.FindElementByName("Preliminar").Click();
            Thread.Sleep(1200);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            RootSession = ReloadSession(RootSession, "TFrmMnSoRCAus");
            var btnTDVI2 = RootSession.FindElementsByClassName("TRadioGroup");

            Thread.Sleep(1000);

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 48, 19);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1000);
            // Flecha hacia bajo en la ventana previa del preli.
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            newFunctions_4.ScreenshotNuevo("Preliminar 4", true, file);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2300);

        }

        static public void ClickBtnPreli5(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            desktopSession.FindElementByName("Preliminar").Click();
            Thread.Sleep(1200);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            RootSession = ReloadSession(RootSession, "TFrmMnSoRCAus");
            var btnTDVI2 = RootSession.FindElementsByClassName("TRadioGroup");

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 48, 19);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1000);

            // Flecha hacia bajo en la ventana previa del preli.
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            newFunctions_4.ScreenshotNuevo("Preliminar 5", true, file);
            Thread.Sleep(900);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2300);

        }


    }
}