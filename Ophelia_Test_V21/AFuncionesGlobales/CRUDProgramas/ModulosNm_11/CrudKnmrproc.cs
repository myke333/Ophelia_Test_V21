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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
    class CrudKnmrproc : FuncionesVitales
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

        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var edit = desktopSession.FindElementsByClassName("TScrollBox");
            //var check1 = desktopSession.FindElementsByName("Activo");
            //var check2 = desktopSession.FindElementsByName("Cargo");
            ////Debugger.Launch();
            


            if (bandera == 0)
            {
                //EDITAR FECHA1
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 225, 30);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

               

                //desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                //Thread.Sleep(1000);

            }
            else
            {
                //Preli
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 835, 312);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(22000);

                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(2000);
            }
        }

        //static public void CrudRproc(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal, int tipo, string file)
        //{

        //    var Elementlist = desktopSession.FindElementsByClassName("TScrollBox");
        //    //var Elementlist2 = desktopSession.FindElementsByClassName("TMonthCalendar");

        //    switch (tipo)
        //    {
        //        case (0):

        //            //EDITAR FECHA1
        //            desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 225, 25);
        //            desktopSession.Mouse.Click(null);
        //            Thread.Sleep(1500);
        //            desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
        //            Thread.Sleep(1000);

        //            //EDITAR FECHA1
        //            desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 480, 25);
        //            desktopSession.Mouse.Click(null);
        //            Thread.Sleep(1500);

        //            desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
        //            Thread.Sleep(1000);


        //            break;




        //    }

            /*

            var cont = 0;

            foreach (var item in Elementlist)
            {
                if (cont == 3)
                {
                    item.SendKeys("2");
                }
                else
                {
                    item.SendKeys(Convert.ToString(cont));
                }

                cont = cont + 1;
                Thread.Sleep(1000);
            }
            */


        
    }
}
