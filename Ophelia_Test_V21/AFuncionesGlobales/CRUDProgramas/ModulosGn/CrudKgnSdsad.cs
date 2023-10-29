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
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{

    class CrudKgnSdsad : FuncionesVitales
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

        static public void CrudSdsad(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal, int tipo, string file)
        {

            var Elementlist = desktopSession.FindElementsByClassName("TScrollBox");


            switch (tipo)
            {
                case (0):

                    //desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 182, 7);
                    //desktopSession.Mouse.DoubleClick(null);
                    //Thread.Sleep(1500);

                    Elementlist[1].SendKeys(crudPrincipal[0]);


                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 82, 40);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);

                    break;




            }

            

            //int cont = 0;

            //foreach (var item in Elementlist)
            //{
            //    if (cont == 3)
            //    {
            //        item.SendKeys("2");
            //    }
            //    else
            //    {
            //        item.SendKeys(Convert.ToString(cont));
            //    }

            //    cont = cont + 1;
            //    Thread.Sleep(1000);
            //}
            


        }
    }
}