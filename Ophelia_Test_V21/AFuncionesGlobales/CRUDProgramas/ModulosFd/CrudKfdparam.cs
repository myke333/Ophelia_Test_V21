using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;



namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdparam : FuncionesVitales
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


        static public void AgregarCrudPri(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {                
                // -------ingresa Pespectiva en Lupa1 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 681, 47);
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

                RootSession1.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // -------ingresa Obj Estrategico en Lupa2 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 687, 92);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // -------ingresa Proyecto en Lupa3 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 687, 157);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);
                                
            }
            else
            {

                // -------ingresa Pespectiva en Lupa1 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 681, 47);
                desktopSession.Mouse.Click(null);

                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                Thread.Sleep(1000);

                RootSession1.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

            }
        }



        static public void ChoiceVentana(WindowsDriver<WindowsElement> desktopSession, int ventana)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (ventana)
            {
                case 1:

                    // Click en ventana Asignaturas/Seciones
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 180, 15);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 2:

                    // Click en ventana Cargos Autorizados
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 320, 15);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 3:

                    // Click en ventana Competencia Adquiridas
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 440, 15);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 4:

                    // Click en ventana Competencia requeridad
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 580, 15);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 5:

                    // Click en ventana Proceso
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 673, 15);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;

                case 6:

                    // Click en ventana Prerequisitos
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 743, 19);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;


                case 7:

                    // Click en la primera ventana
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 26, 19);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;



            }

        }



        static public void AggCrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet123)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {

                case 0:
                    // clic en add registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 178, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 1", true, file);


                    // click en opciones del registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 119, 80);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession1 = null;
                    RootSession1 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession1.Keyboard.SendKeys(crudDet123[0]);
                    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);


                    // click en aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 281, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 1 ", true, file);
                    break;

            

                case 1:  // actualizar y cancelar
                       
                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 247, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE 1 ", true, file);


                    // click en el campo 
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 119, 80);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);


                    RootSession2.Keyboard.SendKeys(crudDet123[1]);
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);


                    // click en Cancelar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 305, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 1 ", true, file);

                    break;

                ////-------------------------------------------------------------------



                case 2:
                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 247, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE 1 ", true, file);


                    // click en el campo 
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 119, 80);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession3 = null;
                    RootSession3 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession3.Keyboard.SendKeys(crudDet123[1]);
                    Thread.Sleep(1000);
                    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 1 ", true, file);
                    Thread.Sleep(1000);


                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 281, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 1 ", true, file);
                    break;

                case 3:

                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 211, 217);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession6 = null;
                    RootSession6 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 1 ", true, file);

                    break;

            }
        }





        static public void AggCrudDet2(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet123)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            switch (bandera)
            {
                case 0:
                    // clic en add registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 178, 425);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 2 ", true, file);


                    // click en opciones del registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 124, 289);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession1 = null;
                    RootSession1 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession1.Keyboard.SendKeys(crudDet123[0]);
                    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);


                    // click en aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 281, 425);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 2 ", true, file);
                    break;



                case 1:
                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 247, 425);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE 2 ", true, file);


                    // click en el campo 
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 124, 289);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    

                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession2.Keyboard.SendKeys(crudDet123[1]);
                    Thread.Sleep(1000);
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 2 ", true, file);
                    Thread.Sleep(1000);


                    // click en Cancelar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 305, 425);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 2", true, file);
                    break;

                // actualizar y aplicar
                case 2:
                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 247, 425);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE 2 ", true, file);


                    // click en el campo 
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 124, 289);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession3 = null;
                    RootSession3 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession3.Keyboard.SendKeys(crudDet123[1]);
                    Thread.Sleep(1000);
                    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 2 ", true, file);
                    Thread.Sleep(1000);


                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 281, 425);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 2 ", true, file);
                    break;


                case 3:
                    // Btn Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 211, 422);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession6 = null;
                    RootSession6 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 2 ", true, file);

                    break;

            }
        }







        static public void AggCrudDet3(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet123)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 0:
                    // clic en add registro. -------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 170, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 3 ", true, file);


                    // click en opciones del registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 110, 69);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession1 = null;
                    RootSession1 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession1.Keyboard.SendKeys(crudDet123[0]);
                    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // click en aplicar ----------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 3 ", true, file);
                    break;

            
                case 1: // actualizar y cancelar
            

                    // clic en Actualizar registro. ---------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  3 ", true, file);


                    // click en el campo 
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 112, 69);
                    desktopSession.Mouse.Click(null);
                    

                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession2.Keyboard.SendKeys(crudDet123[1]);
                    Thread.Sleep(1000);
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 3 ", true, file);
                    Thread.Sleep(1000);


                    // click en Cancelar -----------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 289, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 3 ", true, file);
                    break;


                case 2:  // actualizar y aplicar
                    
                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  3 ", true, file);


                    // click en el campo 
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 112, 69);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    

                    WindowsDriver<WindowsElement> RootSession3 = null;
                    RootSession3 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession3.Keyboard.SendKeys(crudDet123[1]);
                    Thread.Sleep(1000);
                    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 3", true, file);
                    Thread.Sleep(1000);


                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 3", true, file);
                    break;


                case 3:
                    // Btn Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession6 = null;
                    RootSession6 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 3 ", true, file);

                    break;
            }
        }



        static public void AggCrudDet4(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet456)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 0:

                    // clic en BTM add registro. -------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 170, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 4 ", true, file);



                    // click en Tipo Registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 141, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);



                    WindowsDriver<WindowsElement> RootSession1 = null;
                    RootSession1 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession1.Keyboard.SendKeys(crudDet456[0]);
                    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // Click en codigo requisito
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 460, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession2.Keyboard.SendKeys(crudDet456[1]);
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // Click en codigo  especificacion
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 855, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession3 = null;
                    RootSession3 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession3.Keyboard.SendKeys(crudDet456[2]);
                    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);



                    // click en aplicar ----------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 265, 375);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 4 ", true, file);

                    break;

                case 1:


                    // clic en Actualizar registro. ---------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 233, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  4 ", true, file);


                    // Click en codigo especificacion.                
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 845, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    

                    WindowsDriver<WindowsElement> RootSession4 = null;
                    RootSession4 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession4.Keyboard.SendKeys(crudDet456[3]);
                    RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 4 ", true, file);
                    Thread.Sleep(1000);


                    // click en Cancelar -----------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 296, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 4 ", true, file);

                    break;

                case 2:

                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  4 ", true, file);


                    // Click en codigo especificacion.                
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 845, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession5 = null;
                    RootSession5 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession5.Keyboard.SendKeys(crudDet456[3]);
                    RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo(" Datos Insertados CRUD DETALLE 4", true, file);
                    Thread.Sleep(1000);



                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 4", true, file);
                    break;


                case 3:
                    // Btn Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession6 = null;
                    RootSession6 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 4 ", true, file);

                    break;

            }

        }



        static public void AggCrudDet5(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet456)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 0:

                    // clic en BTM add registro. -------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 170, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 5 ", true, file);



                    // click en Tipo Registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 141, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);



                    WindowsDriver<WindowsElement> RootSession1 = null;
                    RootSession1 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession1.Keyboard.SendKeys(crudDet456[0]);
                    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // Click en codigo requisito
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 534, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession2.Keyboard.SendKeys(crudDet456[1]);
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // Click en codigo  especificacion
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 886, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession3 = null;
                    RootSession3 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession3.Keyboard.SendKeys(crudDet456[2]);
                    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);



                    // click en aplicar ----------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 265, 375);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 5 ", true, file);

                    break;

                case 1:


                    // clic en Actualizar registro. ---------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 233, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  5 ", true, file);



                    // Click en codigo especificacion.                
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 879, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession4 = null;
                    RootSession4 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession4.Keyboard.SendKeys(crudDet456[3]);
                    RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 5 ", true, file);
                    Thread.Sleep(1000);


                    // click en Cancelar -----------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 296, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 5 ", true, file);

                    break;

                case 2:

                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  5 ", true, file);


                    ////Click en codigo especificacion.                
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 879, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession5 = null;
                    RootSession5 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession5.Keyboard.SendKeys(crudDet456[3]);
                    RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo(" Datos Insertados CRUD DETALLE 5", true, file);
                    Thread.Sleep(1000);



                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 5", true, file);
                    break;


                case 3:
                    // Btn Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession6 = null;
                    RootSession6 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 5 ", true, file);

                    break;

            }

        }





        static public void AggCrudDet6(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet456)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 0:

                    // clic en BTM add registro. ------------------------------------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 170, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 4 ", true, file);



                    // click en Tipo Registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 104, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    



                    WindowsDriver<WindowsElement> RootSession1 = null;
                    RootSession1 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession1.Keyboard.SendKeys(crudDet456[0]);
                    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // Click en codigo requisito
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 290, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession2.Keyboard.SendKeys(crudDet456[1]);
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);




                    // Click en codigo  especificacion
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 526, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);



                    WindowsDriver<WindowsElement> RootSession3 = null;
                    RootSession3 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession3.Keyboard.SendKeys(crudDet456[2]);
                    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);



                    // click en aplicar ---------------------------------------------------------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 265, 375);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 3 ", true, file);

                    break;

                case 1:


                    // clic en Actualizar registro. ---------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 233, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  3 ", true, file);

                    // Click en codigo especificacion.                
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 855, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);                    

                    WindowsDriver<WindowsElement> RootSession4 = null;
                    RootSession4 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession4.Keyboard.SendKeys(crudDet456[3]);
                    RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 3 ", true, file);
                    Thread.Sleep(1000);


                    // click en Cancelar -----------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 296, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 3 ", true, file);

                    break;

                case 2:

                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  3 ", true, file);


                    // Click en codigo especificacion.                
                    //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 855, 70);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);
                    //desktopSession.Mouse.Click(null);
                    //Thread.Sleep(1000);


                    WindowsDriver<WindowsElement> RootSession5 = null;
                    RootSession5 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    RootSession5.Keyboard.SendKeys(crudDet456[3]);
                    RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 3", true, file);
                    Thread.Sleep(1000);



                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 3", true, file);
                    break;

                case 3:
                    // Btn Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession6 = null;
                    RootSession6 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);

                    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 6 ", true, file);

                    break;

            }

        }






        static public void AggCrudDet7(WindowsDriver<WindowsElement> desktopSession, int bandera, string file, List<string> crudDet456)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 0:

                    // clic en BTM add registro. ------------------------------------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 170, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Insertar Datos CRUD DETALLE 4 ", true, file);


                    // click en Tipo Registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 104, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(crudDet456[0]);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(crudDet456[1]);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(crudDet456[2]);
                    Thread.Sleep(1000);
                    
                    // click en aplicar ---------------------------------------------------------------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 265, 375);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Datos Insertados CRUD DETALLE 3 ", true, file);

                    break;

                case 1:


                    // clic en Actualizar registro. ---------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 233, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  3 ", true, file);

                    // click en Tipo Registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 104, 70);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Keyboard.SendKeys(crudDet456[3]);
                    Thread.Sleep(1000);


                    newFunctions_4.ScreenshotNuevo(" Datos Actualizados CRUD DETALLE 3 ", true, file);
                    Thread.Sleep(1000);


                    // click en Cancelar -----------------
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 296, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Cancelar Datos Editados CRUD DETALLE 3 ", true, file);

                    break;

                case 2:

                    // clic en Actualizar registro.
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    //newFunctions_4.ScreenshotNuevo("Editar CRUD DETALLE  3 ", true, file);


                    desktopSession.Keyboard.SendKeys(crudDet456[3]);
                    Thread.Sleep(1000);


                    newFunctions_4.ScreenshotNuevo(" Datos Actualizados CRUD DETALLE 3 ", true, file);
                    Thread.Sleep(1000);



                    // click en Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Aplicar Datos Editados CRUD DETALLE 3", true, file);
                    break;


                case 3:
                    // Btn Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 378);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> RootSession5 = null;
                    RootSession5 = PruebaCRUD.RootSession();
                    Thread.Sleep(1000);
                    
                    RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);


                    newFunctions_4.ScreenshotNuevo("Datos Eliminados CRUD DETALLE 7 ", true, file);

                    break;

            }

        }






    }
}
