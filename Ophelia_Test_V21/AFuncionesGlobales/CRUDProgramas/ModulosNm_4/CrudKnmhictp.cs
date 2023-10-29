using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmhictp : PruebaCRUD
    {
        public static void CRUDKnmhictp(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                ElementList[5].SendKeys(data[1]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[2]);
            }

        }
        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[0].X + 10, coordenadas[0].Y);
            desktopSession.Mouse.Click(null);
            if (contLupa == 1)
            {
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var ventana = rootSession.FindElementsByClassName("TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                var campo = rootSession.FindElementsByClassName("TEdit");
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(data[0]);
                var btn = rootSession.FindElementsByClassName("TBitBtn");
                foreach (var boton in btn)
                {
                    if (boton.Text == "Aceptar")
                    {
                        rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                }
                indexVal++;
            }

        }
    }
}
