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
	class CrudKSlPrueb:FuncionesVitales
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

        static public void CrudPrueb(WindowsDriver<WindowsElement> desktopsession, string tipo, List<string> crudPrincipal, string file)
		{
            var Elementlist = desktopsession.FindElementsByClassName("TDBEdit");
            var Elementlist2 = desktopsession.FindElementsByClassName("TDBCmbCodigoBox");
            var Elementlist3 = desktopsession.FindElementsByClassName("TDBMemo");
           
                        switch (tipo)
                        {
                            case ("0"):
                    
                                Elementlist[3].SendKeys(crudPrincipal[0]);
                                Elementlist[3].SendKeys(OpenQA.Selenium.Keys.Tab);           
                                Elementlist2[1].SendKeys(crudPrincipal[1]);
                                Elementlist2[1].SendKeys(OpenQA.Selenium.Keys.Tab);
                                Elementlist2[0].SendKeys("Cualitativa");
                                Elementlist[4].SendKeys(crudPrincipal[2]);
                                Elementlist[2].SendKeys(crudPrincipal[3]);
                                Elementlist3[0].SendKeys(crudPrincipal[4]);
                                break;
                            case ("1"):
                                
                                Elementlist2[1].Clear();
                                Elementlist2[1].SendKeys(crudPrincipal[5]);
                                Elementlist2[0].Clear();
                                Elementlist2[0].SendKeys("Cuantitativa");
                               // Elementlist[4].Clear();
                               // Elementlist[4].SendKeys(crudPrincipal[6]);
                                Elementlist[2].Clear();
                                Elementlist[2].SendKeys(crudPrincipal[7]);
                                Elementlist3[0].Clear();
                                Elementlist3[0].SendKeys(crudPrincipal[8]);
                                break;
                        }
               

/*
            
            int cont = 0;

            foreach (var item in Elementlist3)
            {
                if (cont == 8)
                {
                    item.SendKeys("2");
                }
                else
                {
                    item.SendKeys(Convert.ToString(cont));

                }
                cont = cont + 1;

                Thread.Sleep(5000);
            }
     */       
        }
    }
}
