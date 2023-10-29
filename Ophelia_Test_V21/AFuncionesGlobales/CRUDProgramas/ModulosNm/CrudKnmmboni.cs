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
	class CrudKnmmboni : FuncionesVitales
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
			var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
			if (bandera == 0)
			{


				///Click en lupa de identificacion e inserta datos
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 619, 39);
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

				///

				///Click en calendario fecha inicial y enter

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 267, 160);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario Fecha Inicial", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();

				RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



				Thread.Sleep(1000);

				///

				///Click en calendario fecha final y enter

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 372, 160);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario Fecha Final", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();

				RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



				Thread.Sleep(1000);

				///

				///Click en lupa de concepto e inserta datos
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 326, 245);
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

				///

				///Click en lupa de prototipo e inserta datos
				desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 662, 245);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession5 = null;
				RootSession5 = PruebaCRUD.RootSession();



				RootSession5.Keyboard.SendKeys(crudPrincipal[2]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///

				///Click en calendario Fecha Reconocimiento y enter

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 120, 301);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario Fecha Reconocimiento", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession6 = null;
				RootSession6 = PruebaCRUD.RootSession();

				RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



				Thread.Sleep(1000);

				///

				///Click en calendario Fecha pago y enter

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 255, 301);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario Fecha pago", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession7 = null;
				RootSession7 = PruebaCRUD.RootSession();

				RootSession7.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



				Thread.Sleep(1000);

				///

				///Click en campo valor e inserta datos
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 313, 301);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);






			}

			else
			{
				///Click en campo valor y edita datos

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 313, 301);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);


			}



			//Thread.Sleep(1000);




		}




	}
}



