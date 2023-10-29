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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmnovss : FuncionesVitales
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



		static public void AgregarDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
		{

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
			var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
			if (bandera == 0)
			{


				///Lupa identifiacion 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 817, 47);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1055, 43);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 760, 250);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 928, 300);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1075, 60);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession1 = null;
				RootSession1 = PruebaCRUD.RootSession();



				//RootSession1.Keyboard.SendKeys(crudPrincipal[0]);
				//Thread.Sleep(1000);
				//desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				//desktopSession.Mouse.Click(null);
				//Thread.Sleep(1000);
				//WindowsDriver<WindowsElement> RootSession11 = null;
				//RootSession11 = PruebaCRUD.RootSession();
				//RootSession11.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				//Thread.Sleep(2000);



				///

				///Tipo de Novedad
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 75, 113);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 75, 113);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Dias de Vac.Anticipadas

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///IBC. Vacaciones anticipadas

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Indicador Integral

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Valor Ley 1393

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 509, 113);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 509, 113);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cantidad Horas laboradas

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Ord Liquidacion

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Lupa concepto 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 451, 189);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1027, 47);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 758, 259);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1039, 59);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();



				RootSession2.Keyboard.SendKeys(crudPrincipal[2]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///Dias Cotizados Salud

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 308, 307);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				//IBC Salud

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///%Salud Patron 

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				//%Salud Empleado

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cotizacion Salud Anticipada

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cotizacion salud empleado

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Fecha Inicial Novedad en PEriodo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 147, 153);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();

				RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				Thread.Sleep(1000);

				///Fecha Final Nov. Periodo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 291, 153);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();

				RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				Thread.Sleep(1000);



				///F. Incial Nov. en Periodo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 445, 153);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession5 = null;
				RootSession5 = PruebaCRUD.RootSession();

				RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				Thread.Sleep(1000);

				//F. Final Real en Periodo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 581, 153);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession6 = null;
				RootSession6 = PruebaCRUD.RootSession();

				RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				Thread.Sleep(1000);

				///Cambio de Ventana Pension 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 97, 228);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				////Dias Cotizados Pension 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 154, 281);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				//IBC pension 

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				//%Pension Empleado

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///%Pension PAtrono

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cotizacion Pension Empleado

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cotizacion pension Anticipado

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Aportes Voluntarios Empleado

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Aportes Voluntarios Patrono 

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///AporteFondo Solidaridad

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Aporte fondo soli anticipado 

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///% Fondo Subsistencia 

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Aporte Fondo Subsistencia


				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Ap. Fondo subsistencia Antici

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cambio de ventana Riesgos

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 152, 228);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///Dias Arl

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 133, 346);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Dias Cotizados Riesgos

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///IBC riesgos

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Centro de trabajo 

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);


				//Aporte ARL Contratistas

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Cambio de Ventana Parafiscales

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 208, 228);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///IBC Cajas

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 355, 349);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);



				//Dias Cotizados caja comp

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 495, 349);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);





			}

			else
			{
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 355, 349);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);


			}



			//Thread.Sleep(1000);




		}


	}
}
