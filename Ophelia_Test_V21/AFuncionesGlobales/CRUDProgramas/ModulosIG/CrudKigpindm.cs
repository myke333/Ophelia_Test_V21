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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosIG
{
    class CrudKigpindm
    {
        public static void Crud(WindowsDriver<WindowsElement> desktopSession, int Tipo, string file)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox"); ;
            if (Tipo == 1)
            {
                // choice dato
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 41, 143);
                desktopSession.Mouse.Click(null);                
                newFunctions_4.ScreenshotNuevo("Escoge dato", true, file);
                Thread.Sleep(2000);


                // click en pasar
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 775, 198);
                desktopSession.Mouse.Click(null);                
                newFunctions_4.ScreenshotNuevo("Click en pasar", true, file);
                Thread.Sleep(2000);


                // click en guardar 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 712, 390);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Click en guardar", true, file);
                Thread.Sleep(2000);


                
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                newFunctions_4.ScreenshotNuevo("Enteer", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

               
            }
            else if (Tipo == 2)
            {

                // choice dato
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 833, 141);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Escoge dato", true, file);
                Thread.Sleep(2000);
                Thread.Sleep(2000);


                // click en pasar
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 775, 250);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Click en pasar", true, file);
                Thread.Sleep(2000);


                // click en guardar 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 712, 390);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Click en guardar", true, file);
                Thread.Sleep(2000);


                newFunctions_4.ScreenshotNuevo("Se deja como estaba.", true, file);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
        }


    }
}
