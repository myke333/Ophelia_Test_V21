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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
	class CrudKnmresfa : FuncionesVitales
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
			if (bandera == 0)
			{



				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 32, 35);
				desktopSession.Mouse.DoubleClick(null);
				desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 90, 29);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 197, 35);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 288, 35);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
				Thread.Sleep(1000);


				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
				Thread.Sleep(1000);








			}

			else
			{

				var btnTDVI1 = desktopSession.FindElementsByClassName("TScrollBox");

				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 686, 428);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2500);

				newFunctions_4.ScreenshotNuevo("CLICK EN ACEPTAR", true, file);
				Thread.Sleep(3000);



				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 850, 307);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2500);

				newFunctions_4.ScreenshotNuevo("CLICK EN SI", true, file);
				Thread.Sleep(3000);



				WindowsDriver<WindowsElement> RootSession1 = null;
				RootSession1 = PruebaCRUD.RootSession();
				RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				Thread.Sleep(4000);

				newFunctions_4.ScreenshotNuevo("ENTER EN VENTANA AUXILIAR", true, file);
				Thread.Sleep(1000);

				//WindowsDriver<WindowsElement> RootSession2 = null;
				//RootSession2 = PruebaCRUD.RootSession();
				//RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				//Thread.Sleep(1000);
				//newFunctions_4.ScreenshotNuevo("ENTER EN VENTANA AUXILIAR", true, file);
				//Thread.Sleep(2500);
				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 789, 249);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);

                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
				RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				Thread.Sleep(3000);
				newFunctions_4.ScreenshotNuevo("ENTER EN VENTANA AUXILIAR", true, file);

				Thread.Sleep(2000);

				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 1300, 43);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2500);

				newFunctions_4.ScreenshotNuevo("VENTANA AUXILIAR Y EXCEL", true, file);
				Thread.Sleep(2000);
				///ULTIMA VENTANA AUXILIAR

				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 789, 249);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);

				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();

				RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				Thread.Sleep(2000);
				newFunctions_4.ScreenshotNuevo("ENTER EN VENTANA AUXILIAR", true, file);
				Thread.Sleep(2000);






				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 1158, 43);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2500);

				newFunctions_4.ScreenshotNuevo("CIERRE EXCEL GENERADO", true, file);
				Thread.Sleep(2000);

			




				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 891, 433);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				newFunctions_4.ScreenshotNuevo("CLICK EN PRELIMINAR", true, file);
				Thread.Sleep(1000);


			}



			//Thread.Sleep(1000);




		}

		static public void qbe2(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal)
		{




			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

				


			//Coordenadas boton qbe2
			//string filterqb = "13";



			desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 765, 419);
			desktopSession.Mouse.Click(null);
			Thread.Sleep(2000);





			WindowsDriver<WindowsElement> RootSession = null;
			RootSession = PruebaCRUD.RootSession();
			RootSession = ReloadSession(RootSession, "TKQBE");





			var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");





			RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 255, 40);
			RootSession.Mouse.Click(null);
			Thread.Sleep(2000);
			RootSession.Mouse.DoubleClick(null);
			Thread.Sleep(1000);
			RootSession.Keyboard.SendKeys(crudPrincipal[7]);
			Thread.Sleep(1500);
			RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
			Thread.Sleep(2000);



		}



	}
}



