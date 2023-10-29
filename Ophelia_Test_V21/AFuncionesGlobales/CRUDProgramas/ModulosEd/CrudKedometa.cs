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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKedometa
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


        static public void ChoiceVentana(WindowsDriver<WindowsElement> desktopSession, int ventana)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (ventana)
            {
                case 1:

                    // Click en ventana FORMULACION
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 134, 232);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 2:

                    // Click en ventana EVALUACION
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 204, 232);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 3:

                    // Click en ventana REFORMULACION
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 282, 232);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 5:

                    // Click en ventana REFORMULACION
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 232);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

            }

        }


        static public void AggCrudPri(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {                

                // IN NOMBRE EVALUACION
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(500);


                // IN ESTADO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(500);

                // IN ANO 
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);


                // IN Fecha incial
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(500);


                // IN Fecha Final
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(500);

            }
            
            else // Pasos para Editar registro

            {
               
                //  NOMBRE EVALUACION
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                // IN ESTADO
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);


                // IN ANO 
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(500);

            }
        }



        static public void AggCrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");



            if (bandera == 0) // Pasos Obligatorios para agg registro.
            {
                // Click en secuencial
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);


                // in Data Descripcion
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                Thread.Sleep(500);

            }

            else // Pasos para Editar/Actualizar  registro

            {

                // Click en secuencial
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                // Actualizar data Descripcion
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(500);

            }
        }


        static public void AggCrudDet2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");



            if (bandera == 0)
            {

                // Click en Tipo Formulacion
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 131, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);



                // in  Data estado formulacvionRequerimiento 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 233, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);


                // in Fecha Inicio
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                Thread.Sleep(500);


                // in Fecha Final
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(500);



            }

            else // Pasos para Editar/Actualizar  registro

            {                

                // ACTUALIZAR  Fecha Final
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 368, 284);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[2]);
                Thread.Sleep(500);

            }
        }





        static public void AggCrudDet3(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");



            if (bandera == 0) // Pasos Obligatorios para agg registro.
            {

                // Click en PERIODO
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 173, 284);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                //Thread.Sleep(500);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);


                // in  Data TIPO EVALUACION 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 288, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);



                // in  Data ESTADO EVALUACION 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 395, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);



                // in Fecha Inicio
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                Thread.Sleep(500);



                // in Fecha Final
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(500);




            }

            else // Pasos para Editar/Actualizar  registro

            {
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 519, 284);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[2]);
                Thread.Sleep(500);

            }
        }


        static public void AggCrudDet4(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");



            if (bandera == 0)
            {

                // Click en PERIODO
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 177, 284);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                

                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);



                // in Fecha Inicio
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 192, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                Thread.Sleep(500);


                // in Fecha Final
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(500);

            }

            else // Pasos para Editar/Actualizar  registro

            {
                // in Fecha Inicio
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 192, 284);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                //desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                //Thread.Sleep(500);


                // in Fecha Final
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDetalle[2]);
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
