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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo_2
{
	class CrudKsorevac : FuncionesVitales
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


				////Click en lupa de identificacion

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 583, 36);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession1 = null;
				RootSession1 = PruebaCRUD.RootSession();



				RootSession1.Keyboard.SendKeys(crudPrincipal[0]); ///contabilizacion
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				///Click en lupa de vacuna

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 537, 169);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();



				RootSession2.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				////fecha aplicacion 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 208);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();

				RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

				///lupa responsable

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 664, 275);
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
















			}

			else
			{

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 664, 275);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(crudPrincipal[3]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

			}



			//Thread.Sleep(1000);




		}




	}
}


