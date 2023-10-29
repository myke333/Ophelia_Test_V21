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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmjefes : FuncionesVitales
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



        static public void AgregarDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 45, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                
                Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 478, 43);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);

                //Thread.Sleep(1000);




            }

            else
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 45, 73);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2000);
               
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);

                //Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "IEFrame");

                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                //desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                //Thread.Sleep(2000);


            }



            //Thread.Sleep(1000);




        }

        static public void AgregarDatosCrudDetalle(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 143);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                Thread.Sleep(1000);
                









            }

            if (bandera==1)
            {
                //Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 606, 143);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(crudDetalle[3]);
                //Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
            }


            if (bandera==2)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 573, 15);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);



            }
            //Thread.Sleep(1000);




        }

        static public void qbe2(WindowsDriver<WindowsElement> desktopSession)
        {




            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");



            //Coordenadas boton qbe2
            //string filterqb = "13";



            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 726, 344);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);





            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");





            var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");





            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 260, 57);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            RootSession.Mouse.DoubleClick(null);
            Thread.Sleep(1000);
            RootSession.Keyboard.SendKeys("11");
            Thread.Sleep(1500);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);


        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



        }

        //static public void Excel(WindowsDriver<WindowsElement> desktopSession)
        //{
        //    WindowsDriver<WindowsElement> RootSession = null;
        //    RootSession = PruebaCRUD.RootSession();

        //    var btnTDVI3 = RootSession.FindElementsByClassName("");
        //    RootSession.Mouse.MouseMove(btnTDVI3[0].Coordinates, 260, 57);
        //    RootSession.Mouse.Click(null);
        //    Thread.Sleep(2000);

        //    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



        //}

        //static public void AceptarPreliminar(WindowsDriver<WindowsElement> desktopSession)
        //{
        //    Thread.Sleep(1500);
        //    WindowsDriver<WindowsElement> RootSession = null;
        //    RootSession = PruebaCRUD.RootSession();
        //    RootSession = ReloadSession(RootSession, "TFrmInfoP50");

        //    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        //    //var btnTDVI2 = RootSession.FindElementsByClassName("TMessageForm");
        //    //Thread.Sleep(1500);
        //    //RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 254, 92);
        //    //RootSession.Mouse.Click(null);
        //    //Thread.Sleep(1500);


        //}

    }
}


