using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbpmovAc : FuncionesVitales
    {
        public static void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, List<string> CrudPrincipal)
        {
            var TEdit = desktopSession.FindElementsByClassName("TEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 35);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(CrudPrincipal[0]);
            //TEdit[0].SendKeys(CrudPrincipal[0]);
            TEdit[0].SendKeys(CrudPrincipal[1]);
        }

        public static void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, List<string> CrudPrincipal)
        {
            desktopSession.FindElementByName("Aceptar").Click();
                                               
        }

        public static void AgregarCrud3(WindowsDriver<WindowsElement> desktopSession, List<string> CrudPrincipal)
        {
            Thread.Sleep(3000);
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);
                        
        }

        public static void AgregarCrud4(WindowsDriver<WindowsElement> desktopSession, List<string> CrudPrincipal)
        {
            Thread.Sleep(3000);
            WindowsDriver<WindowsElement> RootSession1 = null;
            RootSession1 = PruebaCRUD.RootSession();
            RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);
        }
    }
}
