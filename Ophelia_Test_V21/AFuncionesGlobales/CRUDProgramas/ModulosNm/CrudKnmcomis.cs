using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmcomis : PruebaCRUD
    {
        public static void CRUDKnmcomis(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList= desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                Thread.Sleep(500);
                Thread.Sleep(500);
                LupaDinamica(rootSession, data, 0);
                ElementList[24].SendKeys(data[1]);
                ElementList[21].SendKeys(data[2]);//ya
                ElementList[18].SendKeys(data[3]);//ya
                ElementList[19].Click();//ya
                ElementList[19].SendKeys(data[4]);//ya
                ElementList[35].SendKeys(data[5]);
                ElementList[32].SendKeys(data[6]);
                ElementList[30].SendKeys(data[7]);
                ElementList[28].SendKeys(data[8]);
                Thread.Sleep(500);
                desktopSession.FindElementByName("Autoriza").Click();
                LupaDinamica2(rootSession, data, 1);
                desktopSession.FindElementByName("General").Click();
            }
            else
            {
                ElementList[19].Clear();
                ElementList[19].SendKeys(data[10]);
            }
        }
        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data, int lupa)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = rootSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            if (lupa == 0)
            {
                rootSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[0].X + 10, coordenadas[0].Y);
                rootSession.Mouse.Click(null);
                
            }

            if (contLupa == 1)
            {
                Thread.Sleep(3000);
                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var campo = rootSession.FindElementsByClassName("TEdit");
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                if (lupa == 0)
                {
                    rootSession.Keyboard.SendKeys(data[0]);
                }
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
        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data, int lupa)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            Thread.Sleep(1000);
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(rootSession, 82, 150, 75);
            Thread.Sleep(1000);
            var parentElement = rootSession.FindElement(By.Name(rootSession.Title));
            int contLupa = 1;
            if (lupa == 1)
            {
                rootSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[1].X + 10, coordenadas[1].Y);
                rootSession.Mouse.Click(null);
            }

            if (contLupa == 1)
            {
                Thread.Sleep(2000);

                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                if (lupa == 1)
                {
                    rootSession.Keyboard.SendKeys(data[9]);
                }

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
