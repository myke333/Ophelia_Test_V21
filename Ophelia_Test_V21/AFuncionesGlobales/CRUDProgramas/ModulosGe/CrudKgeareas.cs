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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgeareas /*: FuncionesVitales*/
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

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {

                /* ---------ingresa CodArea en programa*/
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 68, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                // --------ingresa NomArea 
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 303, 31);
                //desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);


                // -------ingresa CodAreaPA
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 627, 31);
                //desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);


            }
            else
            {
                // ingresa dato Actualizado NomAreaAct 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 303, 40);
                Thread.Sleep(500);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);

            }
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 1292, 245);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("Click QBE Det 1", true, file);



            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");

            var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

            RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 16);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);

            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession3 = PruebaCRUD.RootSession();
            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
        }

        // AgregarCrudDet = AgregarCrud1
        static public void AgregarCrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                // ingresa dato EmpresaOrigen,  varMTM = Emplid 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 44, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(1000);


                // ingresa dato lider,  varMTM = lider 
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 147, 26);
                //desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDet1[1]);
                Thread.Sleep(1000);

                
            }
            else
            {
                // ingresa dato EmpresaOrigen,  varMTM = Emplid 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 44, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(1000);


                // ingresa dato lider,  varMTM = lider 
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 147, 26);
                //desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDet1[2]);
                Thread.Sleep(1000);


            }
        }

        static public void Clickk(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 468, 190);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 22, 570);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);


            }
        }
    }
}