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
	class CrudKBiHincr:FuncionesVitales
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

        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                // Datos para add registro
                case 1:

                    // in Cargo
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 32, 45);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                    Thread.Sleep(500);

                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    //  in  PlazasIncrementadas
                    desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                    Thread.Sleep(500);
                    
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    //  in datos Resolucion
                    desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                    Thread.Sleep(500);

                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    //  in Fecha
                    desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                    Thread.Sleep(500);


                    // Click en indicador
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 29, 118);
                    desktopSession.Mouse.Click(null);


                    break;

                // El caso 2 es para  el paso de editar numero de plazas
                case 2:

                    // in Cargo
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 32, 45);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);

                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    //  in  PlazasIncrementadas
                    desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                    Thread.Sleep(500);


                    break;

            }

        }



        //      static public void CrudHincr(WindowsDriver<WindowsElement> desktopSession, int insertar, List<string> CrudPrincipal, string file)
        //{
        //          var selec = PruebaCRUD.NavClass(desktopSession);

        //	switch (insertar)
        //	{
        //              case (0):
        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 250, 115);
        //                  desktopSession.Mouse.Click(null);

        //                  WindowsDriver<WindowsElement> Accep = null;
        //                  Accep = PruebaCRUD.RootSession();
        //                  Accep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(2000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 300, 115);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> Plazas = null;
        //                  Plazas = PruebaCRUD.RootSession();
        //                  Plazas.Keyboard.SendKeys("10");
        //                  Thread.Sleep(2000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 400, 115);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> Resol = null;
        //                  Resol = PruebaCRUD.RootSession();
        //                  Resol.Keyboard.SendKeys("50");
        //                  Thread.Sleep(2000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 550, 115);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> SelFec = null;
        //                  SelFec = PruebaCRUD.RootSession();
        //                  SelFec.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 540, 145);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> Desc = null;
        //                  Desc = PruebaCRUD.RootSession();
        //                  Desc.Keyboard.SendKeys("DescripcionPruebaSQL");
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 95, 185);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 198, 220);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> Accep2 = null;
        //                  Accep2 = PruebaCRUD.RootSession();
        //                  Accep2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 535, 220);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> Accep3 = null;
        //                  Accep3 = PruebaCRUD.RootSession();
        //                  Accep3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, -10, 275);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> Carg = null;
        //                  Carg = PruebaCRUD.RootSession();
        //                  Carg.Keyboard.SendKeys("20");
        //                  break;

        //              case (1):

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 250, 115);
        //                  desktopSession.Mouse.Click(null);

        //                  WindowsDriver<WindowsElement> EditAccep = null;
        //                  EditAccep = PruebaCRUD.RootSession();
        //                  EditAccep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(2000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 300, 115);
        //                  desktopSession.Mouse.DoubleClick(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditPlazas = null;
        //                  EditPlazas = PruebaCRUD.RootSession();
        //                  EditPlazas.Keyboard.SendKeys("20");
        //                  Thread.Sleep(2000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 400, 115);
        //                  desktopSession.Mouse.DoubleClick(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditResol = null;
        //                  EditResol = PruebaCRUD.RootSession();
        //                  EditResol.Keyboard.SendKeys("30");
        //                  Thread.Sleep(2000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 550, 115);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditSelFec = null;
        //                  EditSelFec = PruebaCRUD.RootSession();
        //                  EditSelFec.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 540, 145);
        //                  desktopSession.Mouse.DoubleClick(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditDesc = null;
        //                  EditDesc = PruebaCRUD.RootSession();
        //                  EditDesc.Keyboard.SendKeys("ActualizarDescripcionPruebaSQL");
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 95, 185);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 198, 220);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditAccep2 = null;
        //                  EditAccep2 = PruebaCRUD.RootSession();
        //                  EditAccep2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, 535, 220);
        //                  desktopSession.Mouse.Click(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditAccep3 = null;
        //                  EditAccep3 = PruebaCRUD.RootSession();
        //                  EditAccep3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                  Thread.Sleep(1000);

        //                  desktopSession.Mouse.MouseMove(selec[0].Coordinates, -10, 275);
        //                  desktopSession.Mouse.DoubleClick(null);
        //                  Thread.Sleep(1000);

        //                  WindowsDriver<WindowsElement> EditCarg = null;
        //                  EditCarg = PruebaCRUD.RootSession();
        //                  EditCarg.Keyboard.SendKeys("10");
        //                  break;
        //          }
        //}
    }
}
