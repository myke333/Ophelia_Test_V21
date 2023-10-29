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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
	class CrudKnmncmov : FuncionesVitales
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

				return ApplicationSession;
			}
			catch (Exception ex) { return null; }
		}

		static public void AgregarDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
		{
			
			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
			var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
			if (bandera == 0)
			{


                ///Click en lupa de contabilizacione inserta dato

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 34, 14);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

                //desktopSession.Keyboard.SendKeys("");
                //Thread.Sleep(1000);
                ////desktopSession.Mouse.Click(null);
                ////Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1051, 157);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);



                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 339, 27);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				//----

				//RootSession1.Keyboard.SendKeys(crudPrincipal[0]); ///contabilizacion
    //            Thread.Sleep(1000);
    //            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
    //            desktopSession.Mouse.Click(null);
    //            Thread.Sleep(1000);

                

				///Click en lupa de centro de costo e inserta dato

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 676, 29);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();



				//RootSession2.Keyboard.SendKeys(crudPrincipal[1]); ///centro de costo
				//Thread.Sleep(1000);
				//desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				//desktopSession.Mouse.Click(null);
				//Thread.Sleep(1000);

				///

				///Click en lupa de identificacion proceso e inserta datos

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 447, 79);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();



				RootSession3.Keyboard.SendKeys(crudPrincipal[2]); ///identificacion proceso
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en lupa de concepto e inserta datos 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 348, 176);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(crudPrincipal[3]); ///concepto
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en base e inserta datos
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 387, 176);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[4]); ///base 
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				////

				///Click en valor e inserta dato
				desktopSession.Keyboard.SendKeys(crudPrincipal[5]); ///valor
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				////

				///Click en valor concepto e inserta datos
				desktopSession.Keyboard.SendKeys(crudPrincipal[6]); ///valor concpto acum
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				////

				////Click en Nro.Documento

				desktopSession.Keyboard.SendKeys(crudPrincipal[7]); ///NroDocumento
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				////

				///Click en porcentaje distribucion

				desktopSession.Keyboard.SendKeys(crudPrincipal[8]); ///Porcentaje Distribucion
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				////


				///Click en lupa de Centro de Costo e inserta datos

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 502, 285);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession5 = null;
				RootSession5 = PruebaCRUD.RootSession();



				RootSession5.Keyboard.SendKeys(crudPrincipal[9]); ///centro de costo
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en lupa de tercero a inserta datos

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 502, 327);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession6 = null;
				RootSession6 = PruebaCRUD.RootSession();



				RootSession6.Keyboard.SendKeys(crudPrincipal[10]); ///tercero
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en estado generado
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 562, 324);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Cambio de ventana cuenta contable

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 102, 129);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				////
				///Click en lupa de nombre cuenta contable

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 605, 201);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession7 = null;
				RootSession7 = PruebaCRUD.RootSession();



				RootSession7.Keyboard.SendKeys(crudPrincipal[11]); /// nombre cuenta contable
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en Subcuenta1

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 246);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[12]); ///subcuenta1
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				///

				///Click en subcuenta2

				desktopSession.Keyboard.SendKeys(crudPrincipal[13]); ///subcuenta2
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Click en subcuenta3

				desktopSession.Keyboard.SendKeys(crudPrincipal[14]); ///subcuenta3
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Click en subcuenta4

				desktopSession.Keyboard.SendKeys(crudPrincipal[15]); ///subcuenta4
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				///Click en subcuenta5

				desktopSession.Keyboard.SendKeys(crudPrincipal[16]); ///subcuenta5
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);


				///Click en lupa de clave contable

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 328);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession8 = null;
				RootSession8 = PruebaCRUD.RootSession();



				RootSession8.Keyboard.SendKeys(crudPrincipal[17]); /// clave contable
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en lupa de tipo cuenta por pagar

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 448, 328);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession9 = null;
				RootSession9 = PruebaCRUD.RootSession();



				RootSession9.Keyboard.SendKeys(crudPrincipal[18]); /// TipoCuentaPorPagar
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1062, 55);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

			}

			else
			{
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 246);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[19]); ///subcuenta1
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);




			


			}



			//Thread.Sleep(1000);




		}




	}
}



