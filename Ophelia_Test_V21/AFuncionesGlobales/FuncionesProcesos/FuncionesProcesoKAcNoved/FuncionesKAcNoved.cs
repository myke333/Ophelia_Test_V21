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

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKAcNoved
{
	class FuncionesKAcNoved : FuncionesVitales
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

		static public void CrearVacante(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> ListaCrearVacante, string file)
		{

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
			if (bandera == 0)
			{
				///click en la pestaña movimiento para escoger crear vacante

				newFunctions_4.ScreenshotNuevo("MOVIMIENTO CREAR VACANTE", true, file);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 625, 50);
				desktopSession.Mouse.Click(null);

				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 260, 83);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("la pestaña movimiento se despliega y se escoge CREAR VACANTE", true, file);
				Thread.Sleep(1000);



				///

				///Click en dependencia

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 96, 435);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Click en Dependencia", true, file);
				Thread.Sleep(1000);

				///


				///Se escoge la primera opcion de Estructura Organizacional

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 957, 119);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 882, 148);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Se despliega ventana Arbol y se selecciona la primera opcion de Estructura Organizacional", true, file);
				Thread.Sleep(1000);
				

				///



				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 819, 217);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 819, 233);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 819, 250);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 819, 267);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);
			
				newFunctions_4.ScreenshotNuevo("SELECCION VENTAS", true, file);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 802, 461);
				desktopSession.Mouse.DoubleClick(null);
				Thread.Sleep(1000);

				newFunctions_4.ScreenshotNuevo("Despues de click en aceptar", true, file);
				Thread.Sleep(1000);
				///

				///posterior de aceptar se da click al chulo verde de aceptar

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 775, 513);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Despues de click en chulo verde aceptar", true, file);
				Thread.Sleep(1000);
				///

				///luego se diligencian datos
				///se selecciona Gastos Admon

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 540, 118);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				//click en lupa tipo planta

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 176);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///agregar tipo planta y click  en aceptar de lupa tipo planta
				///
				WindowsDriver<WindowsElement> RootSession = null;
				RootSession = PruebaCRUD.RootSession();
				


				RootSession.Keyboard.SendKeys(ListaCrearVacante[0]); ///tipo planta	=1
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);



				//click en lupa de tipo vinculacion


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1105, 208);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				//agrega tipo vinculacion y click en aceptar de lupa tipo vinculacion

				WindowsDriver<WindowsElement> RootSession1 = null;
				RootSession1 = PruebaCRUD.RootSession();



				RootSession1.Keyboard.SendKeys(ListaCrearVacante[1]);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);


				//click en lupa de cargo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1105, 142);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				//

				////agrega cargo y click en aceptar de lupa cargo

				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();



				RootSession2.Keyboard.SendKeys(ListaCrearVacante[2]); 
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);




				//click en ventana tipo cargo para escoger CARRERA


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1108, 289);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1013, 307);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				///


				///acto administrativo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 850, 331);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();



				RootSession3.Keyboard.SendKeys(ListaCrearVacante[3]);
				Thread.Sleep(2000);
				RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);


				///


				///click en lupa centro de costo
				///

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 769, 292);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				///
				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(ListaCrearVacante[4]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
			


				Thread.Sleep(1000);
                ///

                ///fecha acto administrativo
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1050, 331);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);


                //WindowsDriver<WindowsElement> RootSession5 = null;
                //RootSession5 = PruebaCRUD.RootSession();



                //RootSession5.Keyboard.SendKeys(ListaCrearVacante[5]);
                //Thread.Sleep(2000);
                //RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);



                //Thread.Sleep(1000);


                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1090, 326);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession5 = null;
                RootSession5 = PruebaCRUD.RootSession();



                //RootSession5.Keyboard.SendKeys(ListaCrearVacante[5]);
                //Thread.Sleep(2000);
                RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



                Thread.Sleep(1000);


                ///

                ///pantallazo datos insertados 

                Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Datos insertados", true, file);
				Thread.Sleep(1000);

				///

				//click en aceptar ventana auxiliar con datos ya insertados


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 975, 419);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				newFunctions_4.ScreenshotNuevo("Despues de aceptar datos insertados", true, file);
				Thread.Sleep(3000);

				//

				///click en yes de ventana auxiliar

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 741, 269);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(3000);
				newFunctions_4.ScreenshotNuevo("Click en yes", true, file);
				Thread.Sleep(2000);

                ///

                ///click en ok de ventana satisfactoria de movimiento creado

                WindowsDriver<WindowsElement> RootSession7 = null;
                RootSession7 = PruebaCRUD.RootSession();
				RootSession7.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 873, 172);
                //desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                //newFunctions_4.ScreenshotNuevo("despues de click en ok de ventana satisfactoria de movimiento creado y final del primer movimiento", true, file);
                //Thread.Sleep(1000);


                //Thread.Sleep(1000);
                //RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("despues de click en ok de ventana satisfactoria de movimiento creado y final del primer movimiento", true, file);
				//Thread.Sleep(1000);

				//





			}


        }


		static public void ReservarCargo(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> ListaReservarCargo, string file)
		{

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
			if (bandera == 0)
			{
				///click en la pestaña movimiento para escoger Movimientos Personal Con Reserva

				newFunctions_4.ScreenshotNuevo("MOVIMIENTO RESERVAR CARGO", true, file);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 625, 50);
				desktopSession.Mouse.Click(null);

				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 260, 122);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("la pestaña movimiento se despliega y se escoge Movimientos Personal con Reserva", true, file);
				Thread.Sleep(1000);



				///

				///Se selecciona reserva del cargo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 82, 100);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Se selecciona opcion Reserva del cargo", true, file);

				///



				///Se inserta el campo identificacion e Id planta utilizando la lupa
				//click en lupa de identificacion e id planta para agregar idplanta y click en aceptar


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 200, 294);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1068, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 798, 314);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);



				WindowsDriver<WindowsElement> RootSession10 = null;
				RootSession10 = PruebaCRUD.RootSession();



				RootSession10.Keyboard.SendKeys(ListaReservarCargo[0]);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				newFunctions_4.ScreenshotNuevo("Despues de insertar ID Planta", true, file);
				Thread.Sleep(2000);
				///



				///se da click al chulo verde de aceptar
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 775, 513);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(15000);
				newFunctions_4.ScreenshotNuevo("Despues de click en chulo verde aceptar", true, file);
				Thread.Sleep(1000);

				///

				///luego se diligencian datos en ventana auxiliar

				//click en la lupa de identificacion de empleado para la vacante para ingresar Identificacion
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1018, 224);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1068, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 797, 312);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession11 = null;
				RootSession11 = PruebaCRUD.RootSession();



				RootSession11.Keyboard.SendKeys(ListaReservarCargo[1]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				////



				//click en la lupa de identificacion de empleado que aprueba para ingresar Identificacion

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1018, 301);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);



				WindowsDriver<WindowsElement> RootSession12 = null;
				RootSession12 = PruebaCRUD.RootSession();



				RootSession12.Keyboard.SendKeys(ListaReservarCargo[2]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);





				////


				///insertar Nro. Resolucion

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 691, 427);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				WindowsDriver<WindowsElement> RootSession13 = null;
				RootSession13 = PruebaCRUD.RootSession();



				RootSession13.Keyboard.SendKeys(ListaReservarCargo[3]);
				Thread.Sleep(2000);
				RootSession13.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1500);

				///


				////insertar fecha 


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 850, 422);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession14 = null;
				RootSession14 = PruebaCRUD.RootSession();

				RootSession14.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Datos insertados en Ventana Auxiliar", true, file);
				Thread.Sleep(1000);


				///

				///click en aceptar

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 875, 509);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(15000);
				newFunctions_4.ScreenshotNuevo("Despues de Click en aceptar", true, file);
				Thread.Sleep(3000);



				///click en ok a ventana de movimiento realizado exitosamente

				WindowsDriver<WindowsElement> RootSession15 = null;
				RootSession15 = PruebaCRUD.RootSession();
				RootSession15.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("despues de click en ok de ventana satisfactoria de movimiento creado y final del segundo movimiento", true, file);
				Thread.Sleep(1000);



				///






			}


		}




		static public void CambioCargo(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> ListaCambioCargo, string file)
		{

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
			var btnTDVI2 = desktopSession.FindElementsByClassName("TPanel");
			if (bandera == 0)
			{
				///click en la pestaña movimiento para escoger Movimientos Personal Con Reserva

				newFunctions_4.ScreenshotNuevo("MOVIMIENTO CAMBIO DE CARGO", true, file);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 625, 50);
				desktopSession.Mouse.Click(null);

				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 260, 122);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("la pestaña movimiento se despliega y se escoge Movimientos Personal con Reserva", true, file);
				Thread.Sleep(1000);



				///

				///Se selecciona cambio de cargo

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 379, 75);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Se selecciona opcion cambio de cargo", true, file);

				///

				///Se selecciona opcion nombramiento

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 56, 162);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Se selecciona opcion nombramiento", true, file);

				///

				///Se inserta el campo identificacion e Id planta utilizando la lupa
				//click en lupa de identificacion e id planta para agregar idplanta y click en aceptar


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 200, 294);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1068, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 798, 314);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);



				WindowsDriver<WindowsElement> RootSession20 = null;
				RootSession20 = PruebaCRUD.RootSession();



				RootSession20.Keyboard.SendKeys(ListaCambioCargo[0]);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
				newFunctions_4.ScreenshotNuevo("Despues de insertar ID Planta", true, file);
				Thread.Sleep(2000);
				///



				///se da click al chulo verde de aceptar
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 775, 513);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(7000);
				newFunctions_4.ScreenshotNuevo("Despues de click en chulo verde aceptar", true, file);
				Thread.Sleep(1000);

				///


				///insertar Nro. Resolucion

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 560, 540);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				WindowsDriver<WindowsElement> RootSession21 = null;
				RootSession21 = PruebaCRUD.RootSession();



				RootSession21.Keyboard.SendKeys(ListaCambioCargo[1]);
				Thread.Sleep(2000);
				RootSession21.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
				Thread.Sleep(1500);

				///

				////insertar fecha 


				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 709, 538);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("Ventana Calendario", true, file);
				Thread.Sleep(1000);

				WindowsDriver<WindowsElement> RootSession22 = null;
				RootSession22 = PruebaCRUD.RootSession();

				RootSession22.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



				Thread.Sleep(1000);

				/// 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1036, 544);
				desktopSession.Mouse.Click(null);
				
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);



				WindowsDriver<WindowsElement> RootSession24 = null;
				RootSession24 = PruebaCRUD.RootSession();



				RootSession24.Keyboard.SendKeys(ListaCambioCargo[2]);
				Thread.Sleep(2000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(2000);
			
				newFunctions_4.ScreenshotNuevo("Datos insertados en Ventana Auxiliar", true, file);
				Thread.Sleep(1000);


				///

				///click en aceptar
				//WindowsDriver<WindowsElement> RootSession25 = null;
				//RootSession25 = PruebaCRUD.RootSession();

           
                //Thread.Sleep(1500);
				//RootSession25.Mouse.MouseMove(btnTDVI2[0].Coordinates, 890, 634);
				//RootSession25.Mouse.Click(null);
				desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 890, 634);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(5000);
				newFunctions_4.ScreenshotNuevo("Despues de aceptar", true, file);
				Thread.Sleep(2000);
				//


				///click en ok a ventana de movimiento realizado exitosamente

				WindowsDriver<WindowsElement> RootSession26 = null;
				RootSession26 = PruebaCRUD.RootSession();
				RootSession26.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
				Thread.Sleep(1000);
				newFunctions_4.ScreenshotNuevo("despues de click en ok de ventana satisfactoria de movimiento creado y final del tercer movimiento", true, file);
				Thread.Sleep(1000);



				///


			}


		}





	}
}




