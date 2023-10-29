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
	class CrudKnmrersw : FuncionesVitales
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
				///Id ,Conc ,FecNom ,FecPag ,FechCont ,MotRec ,DocSop ,ValRec ,ValRecAct 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 29);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 301, 29);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 156, 86);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
				Thread.Sleep(1000);


				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);
				desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
				Thread.Sleep(1000);



				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
				Thread.Sleep(1000);


				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
				Thread.Sleep(1000);


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 681, 406);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 651, 438);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


			}

			if (bandera == 1)
			{

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 69, 446);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);

				desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
				Thread.Sleep(1000);


			}
			if (bandera == 2)
			{
				///CLICK EN PRELIMINAR
				var btnTDVI1 = desktopSession.FindElementsByClassName("TGridPanel");
				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 127, 12);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN PRELIMINAR", true, file);
				Thread.Sleep(2000);

				///CLICK EN ACEPTAR PRELIMINAR
				var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 972, 546);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN ACEPTAR", true, file);





			}

			if (bandera == 3)
			{
				///CLICK EN PRELIMINAR
				var btnTDVI1 = desktopSession.FindElementsByClassName("TGridPanel");
				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 127, 12);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN PRELIMINAR", true, file);
				Thread.Sleep(2000);


				///CLICK EN OTRA OPCION DE PRELIMINAR
				var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 972, 173);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN OTRA OPCION DE PRELIMINAR", true, file);

				///CLICK EN ACEPTAR PRELIMINAR
				var btnTDVI3 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI3[0].Coordinates, 972, 546);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN ACEPTAR", true, file);



			}

			if (bandera == 4)
			{
				///CLICK EN PRELIMINAR
				var btnTDVI1 = desktopSession.FindElementsByClassName("TGridPanel");
				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 127, 12);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN PRELIMINAR", true, file);
				Thread.Sleep(2000);

				///CLICK EN OTRA OPCION DE PRELIMINAR
				var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 770, 270);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN OTRA OPCION DE PRELIMINAR", true, file);





				///CLICK EN ACEPTAR PRELIMINAR
				var btnTDVI3 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI3[0].Coordinates, 972, 546);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN ACEPTAR", true, file);





			}

			if (bandera == 5)
			{
				///CLICK EN PRELIMINAR
				var btnTDVI1 = desktopSession.FindElementsByClassName("TGridPanel");
				desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 127, 12);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN PRELIMINAR", true, file);
				Thread.Sleep(2000);

				///CLICK EN OTRA OPCION DE PRELIMINAR
				var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 610, 270);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN OTRA OPCION DE PRELIMINAR", true, file);





				///CLICK EN ACEPTAR PRELIMINAR
				var btnTDVI3 = desktopSession.FindElementsByClassName("TPanel");
				desktopSession.Mouse.MouseMove(btnTDVI3[0].Coordinates, 972, 546);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("CLICK EN ACEPTAR", true, file);





			}



			//Thread.Sleep(1000);




		}




	}
}



