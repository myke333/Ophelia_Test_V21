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
	class CrudKbiparbi:FuncionesVitales
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




        static public void ChoiceVentana(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            // Click en ventana q tiene CrudDetalle
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 198, 11);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

        }


        static public void AggCrudPri(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {

                // Pasos Obligatorios para agg registro.
                // Click Area de Experiencia

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 68);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

                // in Nombre 
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(500);



            }
            // Pasos para Editar registro
            else
            {


                // Click Area de Experiencia
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 68);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                //desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                //Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);

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

                // Click en registro
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 137);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                // in Requerimiento 
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(500);


               



            }
            // Pasos para Editar/Actualizar  registro
            else
            {
                // Pasos Obligatorios para agg registro.

                // Click en registro
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 137);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                //Thread.Sleep(500);


                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                // in Requerimiento 
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








  //      static public void CrudParbi(WindowsDriver<WindowsElement> desktopSession, string tipo, List<string> crudPrincipal, string file)
		//{
  //          Dictionary<string, Point> botonesNavegador = new Dictionary<string, Point>();
  //          botonesNavegador = PruebaCRUD.findname(desktopSession, 30, 0);
  //          var Cambios = PruebaCRUD.NavClass(desktopSession);
            
  //          var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
  //          var Editar = desktopSession.FindElementsByClassName("KBiActem");

  //          WindowsDriver<WindowsElement> Campos1 = null;
  //          Campos1 = PruebaCRUD.RootSession();
  //          WindowsDriver<WindowsElement> Campos2 = null;
  //          Campos2 = PruebaCRUD.RootSession();
  //          WindowsDriver<WindowsElement> Campos3 = null;
  //          Campos3 = PruebaCRUD.RootSession();
  //          WindowsDriver<WindowsElement> Campos4 = null;
  //          Campos4 = PruebaCRUD.RootSession();
  //          WindowsDriver<WindowsElement> Campos5 = null;
  //          Campos5 = PruebaCRUD.RootSession();
  //          WindowsDriver<WindowsElement> Campos6 = null;
  //          Campos6 = PruebaCRUD.RootSession();


		//	switch (tipo)
		//	{
  //              case ("0"):
  //                  //Pestaña 1
  //                  //desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 30, -45);

  //                  //Agregar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Insertar Datos Actividad", true, file);
  //                  Elementlist[2].SendKeys(crudPrincipal[0]);
  //                  Elementlist[2].SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos1.Keyboard.SendKeys(crudPrincipal[1]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  //Aplicar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar", true, file);
  //                  Thread.Sleep(1000);
  //                  //Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos1.Keyboard.SendKeys("10");
  //                  Campos1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos1.Keyboard.SendKeys(crudPrincipal[13]);
  //                  newFunctions_4.ScreenshotNuevo("Editar Datos Actividad", true, file);
  //                  //Descartar edicion - Cancelar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);
  //                  //Confirmar Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos1.Keyboard.SendKeys("10");
  //                  Campos1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos1.Keyboard.SendKeys(crudPrincipal[13]);
  //                  newFunctions_4.ScreenshotNuevo("Datos Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Aplicar Cambios
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar Cambios", true, file);
  //                  //Eliminar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
  //                  //Confirmar Eliminar
  //                  Campos1 = PruebaCRUD.RootSession();
  //                  Campos1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);
  //                  break;

  //              case ("1"):

  //                  //Click Pestaña 2                    
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 150, -45);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  //Agregar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Insertar Datos Clases Discapacidad", true, file);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos2.Keyboard.SendKeys(crudPrincipal[2]);
  //                  Campos2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos2.Keyboard.SendKeys(crudPrincipal[3]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  //Aplicar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar", true, file);
  //                  //Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos2.Keyboard.SendKeys("11");
  //                  Campos2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos2.Keyboard.SendKeys(crudPrincipal[14]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  //Descartar edicion - Cancelar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);
  //                  //Confirmar Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos2.Keyboard.SendKeys("11");
  //                  Campos2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos2.Keyboard.SendKeys(crudPrincipal[14]);
  //                  newFunctions_4.ScreenshotNuevo("Datos Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Aplicar Cambios
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar Cambios", true, file);
  //                  //Eliminar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
  //                  //Confirmar Eliminar
  //                  Campos2 = PruebaCRUD.RootSession();
  //                  Campos2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);
  //                  break;

  //              case ("2"):



  //                  //Click Pestaña 3
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 350, -55);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  //Agregar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Insertar Docs Adicionales", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 10);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos3.Keyboard.SendKeys(crudPrincipal[4]);
  //                  Campos3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos3.Keyboard.SendKeys(crudPrincipal[5]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  //Aplicar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar", true, file);
  //                  Thread.Sleep(1000);
  //                  //Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 10);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos3.Keyboard.SendKeys("110");
  //                  Campos3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos3.Keyboard.SendKeys(crudPrincipal[15]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  //Descartar edicion - Cancelar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);
  //                  //Confirmar Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 10);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos3.Keyboard.SendKeys("110");
  //                  Campos3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos3.Keyboard.SendKeys(crudPrincipal[15]);
  //                  newFunctions_4.ScreenshotNuevo("Datos Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Aplicar Cambios
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar Cambios", true, file);
  //                  //Eliminar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
  //                  //Confirmar Eliminar
  //                  Campos3 = PruebaCRUD.RootSession();
  //                  Campos3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);
  //                  break;


  //              case ("3"):

  //                  //Click Pestaña 4
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 600, -60);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);

  //                  //Agregar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Insertar Listado idiomas", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos4.Keyboard.SendKeys(crudPrincipal[6]);
  //                  Campos4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos4.Keyboard.SendKeys(crudPrincipal[7]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar", true, file);
  //                  Thread.Sleep(1000);
  //                  //Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos4.Keyboard.SendKeys("130");
  //                  Campos4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos4.Keyboard.SendKeys(crudPrincipal[16]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  //Descartar edicion - Cancelar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Confirmar Edicion
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos4.Keyboard.SendKeys("130");
  //                  Campos4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos4.Keyboard.SendKeys(crudPrincipal[16]);
  //                  newFunctions_4.ScreenshotNuevo("Datos Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Aplicar Cambios
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar Cambios", true, file);
  //                  //Eliminar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
  //                  //Confirmar Eliminar
  //                  Campos4 = PruebaCRUD.RootSession();
  //                  Campos4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);
  //                  break;

  //              case ("4"):

  //                  //Click Pestaña 5
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 800, -45);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);

  //                  //Agregar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Insertar instituciones y Especialic.", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[8]);
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[9]);
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[10]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar", true, file);
  //                  Thread.Sleep(1000);
  //                  //Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos5.Keyboard.SendKeys("115");
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[17]);
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[18]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  //Descartar edicion - Cancelar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Confirmar Edicion
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos5.Keyboard.SendKeys("115");
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[17]);
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos5.Keyboard.SendKeys(crudPrincipal[18]);
  //                  newFunctions_4.ScreenshotNuevo("Datos Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Aplicar Cambios
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar Cambios", true, file);
  //                  //Eliminar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
  //                  //Confirmar Eliminar
  //                  Campos5 = PruebaCRUD.RootSession();
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);
                    
  //                  break;


   
  //              case ("5"):


  //                  //Click Pestaña 6
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 1000, -70);
  //                  desktopSession.Mouse.Click(null);
  //                  Thread.Sleep(1000);
  //                  //Agregar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Insertar"].X, botonesNavegador["Insertar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Insertar instituciones y Especialic.", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 10, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos6.Keyboard.SendKeys(crudPrincipal[11]);
  //                  Campos6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos6.Keyboard.SendKeys(crudPrincipal[12]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar", true, file);
  //                  Thread.Sleep(1000);
                    
                    
  //                  break;

  //              case ("6"):

  //                  //Editar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, -100, 5);
  //                  desktopSession.Mouse.Click(null);
  //                  Campos6.Keyboard.SendKeys("45");
  //                  Campos6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos6.Keyboard.SendKeys(crudPrincipal[19]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  //Descartar edicion - Cancelar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Cancelar"].X, botonesNavegador["Cancelar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Descartar Edicion", true, file);
  //                  Thread.Sleep(1000);
  //                  //Confirmar Edicion
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Editar"].X, botonesNavegador["Editar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Editar", true, file);
  //                  Thread.Sleep(1000);
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, -100, 5);
  //                  desktopSession.Mouse.DoubleClick(null);
  //                  Campos6.Keyboard.SendKeys("45");
  //                  Campos6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
  //                  Campos6.Keyboard.SendKeys(crudPrincipal[19]);
  //                  newFunctions_4.ScreenshotNuevo("Ingreso de datos", true, file);
  //                  Thread.Sleep(1000);
  //                  break;

  //              case ("7"):

  //                  //Click Pestaña 6
  //                  desktopSession.Mouse.MouseMove(Elementlist[2].Coordinates, 1000, -70);
  //                  desktopSession.Mouse.Click(null);

  //                  //Aplicar Cambios
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Aplicar"].X, botonesNavegador["Aplicar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Aplicar Cambios", true, file);

  //                  //Eliminar
  //                  desktopSession.Mouse.MouseMove(Cambios[0].Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
  //                  desktopSession.Mouse.Click(null);
  //                  newFunctions_4.ScreenshotNuevo("Eliminar Datos", true, file);
  //                  //Confirmar Eliminar
  //                  Campos5 = PruebaCRUD.RootSession();
  //                  Campos5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
  //                  newFunctions_4.ScreenshotNuevo("Confirmar Eliminar Datos", true, file);
  //                  break;

  //          }


        //}

    }
}
