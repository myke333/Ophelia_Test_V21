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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoinhsg : FuncionesVitales
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

       




        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            List<int> lupasOff = new List<int>() { 3 };
            rootSession = RootSession();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementButton = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");


            if (tipo == 1)
            {

                {
                   
                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, crudVariables, lupasOff);
                    ElementList[18].SendKeys(crudVariables[3]);



                }

            }

            else
            {

                ElementList[18].Clear();
                ElementList[18].SendKeys(crudVariables[4]);

            }
        }

        public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            cerrarBorrar(desktopSession);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);
        }
        public static void cerrarBorrar(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "IEFrame");            
            var allFrame1 = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame1[0], 0, 0).MoveByOffset(allFrame1[0].Size.Width / 2, allFrame1[0].Size.Height / 2).DoubleClick().Click().Perform();
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

		static public void AgregarDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
		{

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

			if (bandera == 0)
			{

				//Click en lupa centro de trabajo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 495, 26);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession1 = null;
				RootSession1 = PruebaCRUD.RootSession();



				RootSession1.Keyboard.SendKeys(crudPrincipal[0]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
				Thread.Sleep(1000);

				//click en hora inspeccion 


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 584, 26);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();

				RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				//click en fecha de inspeccion 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 705, 26);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();

				RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				///Lupa tipo de inpeccion 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 342, 65);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
				Thread.Sleep(1000);









			}

			if (bandera == 1)
			{
				///Lupa tipo de inpeccion 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 342, 65);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(crudPrincipal[2]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Datos Insertados", true, file);
				Thread.Sleep(1000);
			}
		}





	}
}
