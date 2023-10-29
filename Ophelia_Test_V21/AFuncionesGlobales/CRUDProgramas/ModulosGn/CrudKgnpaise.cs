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
    class CrudKgnpaise : PruebaCRUD
    {
        public static void CRUDKgnpaise(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                ElementList[7].SendKeys(data[1]);
                ElementList[6].SendKeys(data[2]);
                ElementList[5].SendKeys(data[3]);
                ElementList[4].SendKeys(data[4]);
            }
            else
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(data[5]);
            }
        }
    }
}
