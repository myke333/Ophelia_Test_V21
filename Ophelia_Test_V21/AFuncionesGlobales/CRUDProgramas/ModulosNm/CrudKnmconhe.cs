using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmconhe : PruebaCRUD
    {
        public static void CRUDKnmconhe(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementList1 = desktopSession.FindElementByClassName("TDBCheckBox");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[4].Click();
                ElementList[7].Click();
                if (data[5] == "SQL")
                {
                    LupaDinamica(rootSession, data, 0);
                    ElementList2[3].SendKeys(data[1]);
                    LupaDinamica(rootSession, data, 2);
                }else
                {
                    LupaDinamica(rootSession, data, 0);
                    ElementList2[3].SendKeys(data[1]);
                    LupaDinamica(rootSession, data, 2);
                }
                    
                ElementList1.Click();
            }
            else
            {
                ElementList1.Click();
            }
        }
        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data, int lupa)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int cont = 0;
            if (lupa == 0)
            {   
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[0].X + 10, coordenadas[0].Y);
                desktopSession.Mouse.Click(null);
            }
            else if (lupa == 1)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[1].X + 10, coordenadas[2].Y);
                desktopSession.Mouse.Click(null);
            }
            else
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[2].X + 10, coordenadas[2].Y);
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
                if (lupa == 0)
                {
                    rootSession.Keyboard.SendKeys(data[0]);
                }else if(lupa ==1)
                {
                    rootSession.Keyboard.SendKeys(data[1]);
                }
                else
                {
                    rootSession.Keyboard.SendKeys(data[2]);
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
