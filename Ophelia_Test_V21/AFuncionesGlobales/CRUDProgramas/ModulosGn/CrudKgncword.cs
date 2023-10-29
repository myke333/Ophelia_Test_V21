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

using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgncword : FuncionesVitales
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

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                //consecutivo
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 107, 106);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);


                //forma de programa
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 231, 106);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                //nombre combinacion  
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 472, 106);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);

                //ruta planilla
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 355, 138);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);

                //ruta de datos
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 313, 189);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);

                //campos generados
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 330, 233);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);





            }
            else
            {
                //ruta de datos
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 313, 189);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Clear);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);

                ////acumulados del año
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 587, 160);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Clear);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                //Thread.Sleep(1000);



            }
        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }

    }
}
