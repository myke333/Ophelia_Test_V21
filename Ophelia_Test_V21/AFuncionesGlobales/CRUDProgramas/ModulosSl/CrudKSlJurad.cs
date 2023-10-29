﻿using OpenQA.Selenium.Appium;
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
	class CrudKSlJurad:FuncionesVitales
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

        static public void CrudJurad(WindowsDriver<WindowsElement> desktopsession, string tipo, List<string> crudPrincipal, string file)
		{

            var Elementlist = desktopsession.FindElementsByClassName("TDBEdit");

			switch (tipo)
			{
                case ("0"):
                    Elementlist[6].SendKeys(crudPrincipal[0]);
                    Elementlist[6].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[5].SendKeys(crudPrincipal[1]);
                    break;
                case ("1"):
                    Elementlist[5].Clear();
                    Elementlist[5].SendKeys(crudPrincipal[3]);
                    break;
			}

            /*
            int cont = 0;
            foreach(var item in Elementlist)
			{
                if(cont == 4)
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
