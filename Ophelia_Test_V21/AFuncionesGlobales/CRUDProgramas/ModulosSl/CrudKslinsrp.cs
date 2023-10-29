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
    class CrudKslinsrp : FuncionesVitales
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

        static public void FechaInscripcion(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle1)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 655, 49);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 552, 68);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle1)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 156, 179);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 427, 178);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle1[0]);
                Thread.Sleep(1500);
            }
        }

        static public void Pestañas(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle2)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 107, 124);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 48, 122);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }
        }

        static public void Click1(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 1)//AGREGAR CRUD
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1727, 337);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 2) //BOTON APLICAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1258, 340);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 3) //BOTON EDITAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1227, 340);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 4) //CANCELAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1285, 340);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }

            if (bandera == 5) //ELIMINAR
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1200, 340);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
            }

            if (bandera == 6) //ELIMINAR CRUD PRINCIPAL
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1170, 340);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
            }
        }


        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle2)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 68, 177);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle2[0]);
                Thread.Sleep(1500);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 239, 177);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle2[1]);
                Thread.Sleep(1500);
            }
        }

        static public void Preliminar(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
                var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

                if (bandera == 1) //Preliminar Reporte 1
                {
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 122, -50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession = null;
                    RootSession = PruebaCRUD.RootSession();
                    RootSession = ReloadSession(RootSession, "TFrmSlParRe");

                    var btnTDVI2 = RootSession.FindElementByName("Reporte 1");

                    RootSession.Mouse.MouseMove(btnTDVI2.Coordinates, 5, 5);
                    RootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);


                }

                if (bandera == 2) //Preliminar Reporte 2
                {
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 122, -50);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession = null;
                    RootSession = PruebaCRUD.RootSession();
                    RootSession = ReloadSession(RootSession, "TFrmSlParRe");

                    var btnTDVI2 = RootSession.FindElementByName("Reporte 2");

                    RootSession.Mouse.MouseMove(btnTDVI2.Coordinates, 8, 5);
                    RootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
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