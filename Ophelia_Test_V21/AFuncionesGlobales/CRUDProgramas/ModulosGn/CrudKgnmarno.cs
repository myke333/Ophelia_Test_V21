﻿using System;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnmarno : PruebaCRUD
    {
        public static void CRUDKgnmarno(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data, 0);
                LupaDinamica2(desktopSession, data, 1);
                ElementList[9].SendKeys(data[2]);
                ElementList[14].SendKeys(data[2]);
                ElementList[13].SendKeys(data[3]);
                ElementList[12].SendKeys(data[4]);
                ElementList[17].SendKeys(data[5]);
            }
            else
            {
                ElementList[9].Clear();
                ElementList[9].SendKeys(data[6]);
            }

        }

        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data, int Lupa)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            //foreach (Point coord in coordenadas)
            //{
            Thread.Sleep(2000);

            if (Lupa == 0)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[0].X + 10, coordenadas[0].Y);
                desktopSession.Mouse.Click(null);
            }
            if (Lupa == 1)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[1].X + 10, coordenadas[1].Y);
                desktopSession.Mouse.Click(null);
            }
            if (contLupa == 1)
            {
                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                if(Lupa == 0)
                {
                    rootSession.Keyboard.SendKeys(data[0]);
                }

                if (Lupa == 1)
                {
                    rootSession.Keyboard.SendKeys(data[1]);
                }

                //Aqui se puede enviar un valor

                var btn = rootSession.FindElementsByClassName("TBitBtn");
                ////Debugger.Launch();
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
            //}

        }
    }
}
