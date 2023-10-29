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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmcsmam : FuncionesVitales
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

        //static public void ChoiceVentana(WindowsDriver<WindowsElement> desktopSession, int ventana)
        //{
        //    var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

        //    switch (ventana)
        //    {
        //        case 1:  // Click en ventana Adicional.

        //            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 248, 77);
        //            desktopSession.Mouse.Click(null);
        //            Thread.Sleep(1500);
        //            break;

        //        case 2:  // Click en Ventana Salarios.
                    
        //            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 298, 77);
        //            desktopSession.Mouse.Click(null);
        //            Thread.Sleep(1500);
        //            break;

        //        case 3:  // Click en Ventana firma.

        //            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 342, 77);
        //            desktopSession.Mouse.Click(null);
        //            Thread.Sleep(1500);
        //            break;               
        //    }
        //}


        static public void InData(WindowsDriver<WindowsElement> desktopSession, string file, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            
                // -------ingresa data identi
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 44, 38);
                desktopSession.Mouse.Click(null);

                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);


                // -------ingresa data Departamento- Lupa 1
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, 120);
                desktopSession.Mouse.Click(null);

                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                Thread.Sleep(1000);


                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);

                RootSession1.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // -------In Data Municipio en Lupa2 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 639, 120);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);

                RootSession1.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Datos ingresados en Ventana DATOS DE LA EMPRESA", true, file);

                /////////////////////////////////////////////////////////////////////////////////////////////////

                // -------- Click en ventana Adicional.. 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 248, 77);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Click en ventana Adicional", true, file);

                /////////////////////////////////////////////////////////////////////////////////////////////////


                // -------In Data Ciuda D Expedicion
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 66, 124);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys("100");
                Thread.Sleep(500);

                 newFunctions_4.ScreenshotNuevo("Datos ingresados en Ventana ADICIONAL", true, file);


                /////////////////////////////////////////////////////////////////////////////////////////////////

                // ----------  Click en Ventana Salarios...                
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 298, 77);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Click en Ventana Salarios", true, file);



                /////////////////////////////////////////////////////////////////////////////////////////////////

                // -------In Data ---  Asignacion Basica Mensual --- Lupa4 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 326, 112);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);


                // -------In Data -----  PrimaTecnica --  Lupa5 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 326, 158);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);


                // -------In Data ---   REMUNERACION TRABAJO  ------  Lupa6
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 326, 196);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // -------In Data ---   REMUNERACION SERVICIO  ------  Lupa7
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 326, 234);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // FILA2 DE LUPAS


                // -------In Data ---   GASTOS REPRESENTACION  ------  Lupa8
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 662, 112);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                // -------In Data ---   PRIMA ANTIGUEDAD  ------  Lupa9
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 662, 158);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[11]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                // -------In Data ---   REMUNERACION TRABAJO  ------  Lupa10
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 662, 196);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(crudPrincipal[12]);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                newFunctions_4.ScreenshotNuevo("Datos ingresados en Ventana SALARIOS", true, file);


                /////////////////////////////////////////////////////////////////////////////////////////////////

                // ----------  Click en Ventana Firma ...

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 342, 77);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Click en Ventana Firma", true, file);

                /////////////////////////////////////////////////////////////////////////////////////////////////

                // -------In Data Funcionario
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 66, 119);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[13]);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);                
                desktopSession.Keyboard.SendKeys(crudPrincipal[14]);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[15]);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[16]);
                Thread.Sleep(500);

                newFunctions_4.ScreenshotNuevo("Datos ingresados en Ventana FIRMA", true, file);


            /////////////////////////////////////////////////////////////////////////////////////////////////

            // ----------  Click en TIPO DE FORMATO ...

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 18, 290);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("SE DA CLICK EN TIPO DE FORMATO", true, file);

            /////////////////////////////////////////////////////////////////////////////////////////////////

        }



        static public void ClickBtnAceptar(WindowsDriver<WindowsElement> desktopSession)
        {
            desktopSession.FindElementByName("Aceptar").Click();
            Thread.Sleep(1500);
        }






    }
}
