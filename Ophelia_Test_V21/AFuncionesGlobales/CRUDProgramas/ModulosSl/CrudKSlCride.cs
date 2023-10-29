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
	class CrudKSlCride:FuncionesVitales
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

        static public void CrudCride(WindowsDriver<WindowsElement> desktopsession, string tipo, List<string> CrudPrincipal, string file)
		{
            var Elementlist = desktopsession.FindElementsByClassName("TDBEdit");

			switch (tipo)
			{
                case ("0"):
                    Elementlist[4].SendKeys(CrudPrincipal[0]);
                    Elementlist[3].SendKeys(CrudPrincipal[1]);
                    Elementlist[2].SendKeys(CrudPrincipal[2]);
                    break;
                case ("1"):
                    Elementlist[3].Clear();
                    Elementlist[3].SendKeys(CrudPrincipal[3]);
                    Elementlist[2].Clear();
                    Elementlist[2].SendKeys(CrudPrincipal[4]);
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

                Thread.Sleep(2000);
			}
*/
		}
    }
}
