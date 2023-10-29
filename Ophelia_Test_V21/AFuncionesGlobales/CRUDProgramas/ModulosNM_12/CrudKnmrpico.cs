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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_12
{
	class CrudKnmrpico : FuncionesVitales
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

				//Click en la lupa para luego ingresar Protipo Liquidacion

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 433, 62);
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




			}

			if (bandera == 1)
			{
				//click en aceptar

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 732, 367);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(3000);
				newFunctions_4.ScreenshotNuevo("Despues de click en aceptar", true, file);
				Thread.Sleep(1000);
				///
			}
		}


		static public void qbe2(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal, string file)
		{


			///click en boton qbe

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


			desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 840, 367);
			desktopSession.Mouse.Click(null);
			Thread.Sleep(2000);
			newFunctions_4.ScreenshotNuevo("Despues de click en boton QBE", true, file);
			Thread.Sleep(1000);
			///

			/// Reconoce ventana de qbe 

			WindowsDriver<WindowsElement> RootSession = null;
			RootSession = PruebaCRUD.RootSession();

			var btnTDVI2 = RootSession.FindElementsByClassName("TStringGrid");
			///

			///Ubica campo de identificacion

			RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 263, 39);
			RootSession.Mouse.Click(null);


			Thread.Sleep(2000);

			RootSession.Mouse.DoubleClick(null);
			Thread.Sleep(1000);

			///


			/// introduce filtro qbe y da enter
			RootSession.Keyboard.SendKeys(crudPrincipal[1]);
			Thread.Sleep(1500);
			newFunctions_4.ScreenshotNuevo("Filtro QBE insertado", true, file);
			Thread.Sleep(1500);
			RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
			Thread.Sleep(2000);
			newFunctions_4.ScreenshotNuevo("Despues de aplicar filtro", true, file);
			Thread.Sleep(1000);
			//



		}




	}
}



