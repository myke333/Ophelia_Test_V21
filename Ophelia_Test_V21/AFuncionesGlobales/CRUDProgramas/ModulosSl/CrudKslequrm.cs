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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
    class CrudKslequrm : FuncionesVitales
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

        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 1)//AGREGAR CRUD
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1723, 297);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 2) //BOTON APLICAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1814, 297);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 3) //BOTON EDITAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1779, 297);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 4) //CANCELAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1830, 297);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }

            if (bandera == 5) //ELIMINAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1749, 297);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle1)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 116, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 80, 192);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 313, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle1[0]);
                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 891, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 868, 148);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 923, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle1[1]);
                Thread.Sleep(1500);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 923, 117);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle1[2]);
                Thread.Sleep(1500);
            }
        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle2)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 114, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 73, 136);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 313, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 407, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle2[0]);
                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 728, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 705, 148);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 922, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle2[1]);
                Thread.Sleep(1500);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 497, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle2[2]);
                Thread.Sleep(1500);
            }
        }

        static public void Pestañas(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 170, 68);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 150, 68);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 1292, 245);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1500);

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
    }
}
