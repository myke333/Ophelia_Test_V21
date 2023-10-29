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
    class CrudKbientre : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();
            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 112, 52);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);

                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 212, 52);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 64, 166);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 423, 166);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 493, 166);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 493, 176);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 592, 166);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 749, 166);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 239, 219);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 395, 219);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 73, 271);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 344, 271);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 588, 271);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 51, 309);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(2000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 239, 219);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 588, 271);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(2000);
            }
        }

        static public void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle1)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI3 = desktopSession.FindElementsByClassName("TTabSheet");
            
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();
            if (bandera == 0)
            {
              
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 63, 179);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDetalle1[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                Thread.Sleep(2000);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Space);

                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 598, 179);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDetalle1[1]);

                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 63, 255);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDetalle1[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 579, 179);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudDetalle1[3]);
 
            }
        }

        static public void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle2)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI3 = desktopSession.FindElementsByClassName("TTabSheet");

            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();
            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 83, 167);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                desktopSession.Keyboard.SendKeys(crudDetalle2[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);


                desktopSession.Keyboard.SendKeys(crudDetalle2[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);

                desktopSession.Keyboard.SendKeys(crudDetalle2[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);

                desktopSession.Keyboard.SendKeys(crudDetalle2[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2000);

                desktopSession.Keyboard.SendKeys(crudDetalle2[4]);
                Thread.Sleep(2000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 1274, 167);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudDetalle2[5]);
                Thread.Sleep(2000);
            }
        }


        static public void Cambiar(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI3 = desktopSession.FindElementsByClassName("TTabSheet");

            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 61, 102);
                    desktopSession.Mouse.Click(null);
                    break;

                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 137, 102);
                    desktopSession.Mouse.Click(null);
                    break;

                case 2:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 185, 102);
                    desktopSession.Mouse.Click(null);
                    break;
            }
           
        }

    }
}
