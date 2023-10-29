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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
    class CrudKnmlitss : FuncionesVitales
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
            //for (int i = 0; i < edit.Count; i++)
            //{
            //    edit[i].SendKeys(i.ToString());
            //    Thread.Sleep(2000);

            //}


            if (bandera == 0)
            {
                //TTIPO CONTA
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 80, 80);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);

                //FECHA 1
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 325, 91);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                //FECHA 2
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 455, 91);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                //FECHA 3
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 580, 91);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                //TIPO NOMINA
                desktopSession.Mouse.MouseMove(edit[0].Coordinates, 248, 140);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);

                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);



                //desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                //Thread.Sleep(1000);


                ////EJEMPLOS CRUD
                //desktopSession.Mouse.MouseMove(edit[0].Coordinates, 50, 50);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession1 = null;
                //RootSession1 = PruebaCRUD.RootSession();
                //RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);



            }
            //else
            //{
            //    //EDITAR VALOR VARIABLE
            //    desktopSession.Mouse.MouseMove(edit[0].Coordinates, 537, 113);
            //    desktopSession.Mouse.DoubleClick(null);
            //    Thread.Sleep(1500);
            //    desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
            //    Thread.Sleep(1000);

            //}
        }

        //public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        //{
        //    desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
        //    desktopSession.Mouse.Click(null);
        //    Thread.Sleep(500);
        //    newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
        //    WindowsDriver<WindowsElement> RootSession1 = null;
        //    RootSession1 = PruebaCRUD.RootSession();
        //    RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //    WindowsDriver<WindowsElement> RootSession2 = null;
        //    RootSession2 = PruebaCRUD.RootSession();
        //    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //    newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
        //    Thread.Sleep(2000);
        //}

        //static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        //{
        //    WindowsDriver<WindowsElement> RootSession2 = null;
        //    RootSession2 = PruebaCRUD.RootSession();
        //    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //    Thread.Sleep(2000);

        //}

    }
}
