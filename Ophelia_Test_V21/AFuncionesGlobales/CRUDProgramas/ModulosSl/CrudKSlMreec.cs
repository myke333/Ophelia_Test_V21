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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
	class CrudKSlMreec:FuncionesVitales
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

        static public void CrudMreec(WindowsDriver<WindowsElement> desktopSession, string tipo, List<string>CrudPrincipal, string file)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");

			switch (tipo)
			{
                case ("0"):
                    Elementlist[2].SendKeys(CrudPrincipal[1]);
                    Elementlist[3].SendKeys(CrudPrincipal[0]);
                    Elementlist[5].SendKeys(CrudPrincipal[2]);
                    break;
                case ("1"):
                    Elementlist[2].Clear();
                    Elementlist[2].SendKeys(CrudPrincipal[4]);
                    Elementlist[3].Clear();
                    Elementlist[3].SendKeys(CrudPrincipal[3]);
                    Elementlist[5].Clear();
                    Elementlist[5].SendKeys(CrudPrincipal[5]);
                    break;
			}
/*
            int cont = 0;

            foreach(var item in Elementlist)
			{
                if(cont == 3)
				{
                    item.SendKeys("2");
				}
				else
				{
                    item.SendKeys(Convert.ToString(cont));
				}

                cont = cont + 1;
			}
*/
		}

        static public void BuscarRegistro(WindowsDriver<WindowsElement> desktopSession, int proceso)
		{
            var Buscar = PruebaCRUD.NavClass(desktopSession);

			switch (proceso)
			{
                case (0):
                    //Primer Lupa
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 18, 130);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    //Boton Aceptar Filtro
                    WindowsDriver<WindowsElement> btnAcep = null;
                    btnAcep = PruebaCRUD.RootSession();
                    btnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
                    //Segunda Lupa
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 25, 200);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    //Ingresar Filtro
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 750, 275);
                    desktopSession.Mouse.Click(null);
                    WindowsDriver<WindowsElement> NumFiltro = null;
                    NumFiltro = PruebaCRUD.RootSession();
                    NumFiltro.Keyboard.SendKeys("76");
                    Thread.Sleep(2000);
                    //Aceptar Primer Filtro
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 1100, 265);
                    desktopSession.Mouse.Click(null);
                    //Aceptar Segundo Filtro
                    desktopSession.Mouse.Click(null);

                    break;
                case (1):
                    //Primer Lupa
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 18, 130);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    //Boton Aceptar Filtro
                    WindowsDriver<WindowsElement> EditbtnAcep = null;
                    EditbtnAcep = PruebaCRUD.RootSession();
                    EditbtnAcep.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
                    //Segunda Lupa
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 25, 200);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    //Ingresar Filtro
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 750, 275);
                    desktopSession.Mouse.Click(null);
                    WindowsDriver<WindowsElement> EditNumFiltro = null;
                    EditNumFiltro = PruebaCRUD.RootSession();
                    EditNumFiltro.Keyboard.SendKeys("676");
                    Thread.Sleep(2000);
                    //Aceptar Primer Filtro
                    desktopSession.Mouse.MouseMove(Buscar[0].Coordinates, 1100, 265);
                    desktopSession.Mouse.Click(null);
                    //Aceptar Segundo Filtro
                    desktopSession.Mouse.Click(null);
                    break;
            }

            
        }
    }
}
