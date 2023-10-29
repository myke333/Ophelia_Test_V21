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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsogiate : FuncionesVitales
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



        static public void ChoiceVentana(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            // Click en ventana q tiene CrudDetalle
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 148, 75);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);


            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 244, 130);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);

        }


        static public void AggCrudPri(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)   // Pasos Obligatorios para agg registro.
            {

                // Click en la lupa
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 619, 41);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);

                RootSession2.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2500);


                


            }
            // Pasos para Editar registro
            else
            {

                // Click en fecha inicio y actualizar
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 136);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);


                // in Nombre 
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(500);


                // click en calendario
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 244, 130);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);


            }
        }




        static public void AggCrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {
               
                // Click en registro
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 106, 221);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                
            }
            // Pasos para Editar/Actualizar  registro
            else
            {
                // Click en registro
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 106, 221);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                desktopSession.Keyboard.SendKeys(crudDet1[1]);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

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
