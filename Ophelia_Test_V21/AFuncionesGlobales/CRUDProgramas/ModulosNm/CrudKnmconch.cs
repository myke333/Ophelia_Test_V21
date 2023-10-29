using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmconch : PruebaCRUD
    {
        public static void CRUDKnmconch(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                LupaDinamica1(rootSession, data,0);
                LupaDinamica1(rootSession, data,1);
                ElementList[6].SendKeys(data[2]);
                ElementList[5].SendKeys(data[3]);
                ElementList[4].SendKeys(data[4]);
            }
            else
            {
                ElementList[4].Clear();
                ElementList[4].SendKeys(data[5]);
            }
        }
        public static void LupaDinamica1(WindowsDriver<WindowsElement> desktopSession, List<string> data, int lupa)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            if (lupa == 0)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[0].X + 10, coordenadas[0].Y);
                desktopSession.Mouse.Click(null);
            }
            else
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[1].X + 10, coordenadas[1].Y);
                desktopSession.Mouse.Click(null);
            }

            if (contLupa == 1)
            {
                Thread.Sleep(500);

                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(data[lupa]);

                var btn = rootSession.FindElementsByClassName("TBitBtn");
                foreach (var boton in btn)
                {
                    if (boton.Text == "Aceptar")
                    {
                        rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                }
            }
        }
    }
}
