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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosRl
{
    class CrudKrlednfo
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
                //
                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }


        static public void ChoiceVentana(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            // Click en ventana q tiene CrudDetalle
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 66);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

        }


        static public void AggCrudPri(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                // Pasos Obligatorios para agg registro.

                // Click en Identificacion
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 96, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                // in Modalidad Academica
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(500);

            }
            // Pasos para Actualizar registro
            else
            {

                // Click en Modalidad Academica
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 39, 111);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);                

                // in Nombre 
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(500);

            }
        }


        static public void AggCrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                // Pasos Obligatorios para agg registro.

                // Doble Click en TIPO
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 161, 117);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                // Se escoge el Tipo
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TCaptura");
                Thread.Sleep(1400);


                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(800);

                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(800);

                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(800);

                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(800);

                RootSession.Keyboard.SendKeys(crudDetalle[0]);
                Thread.Sleep(500);

                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(400);

                //////////////////////////////////////////////////////////


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                // ingresa Requisito 
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(500);



                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                // ingresa Especificacion
                desktopSession.Keyboard.SendKeys(crudDetalle[2]);
                Thread.Sleep(500);

            }


            else  // Pasos para Editar/Actualizar  registro
            {
                // Click en TIPO
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 163, 110);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                //Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                // in Requisito 
                //desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                //Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                // Actualizar Especificacion
                desktopSession.Keyboard.SendKeys(crudDetalle[3]);
                Thread.Sleep(500);



            }
        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }

    }
}
