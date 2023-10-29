﻿using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.Drawing;
using OpenQA.Selenium;
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKed3rela
    {
        public static void KEd3relaCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);
            }
            else {
                List<int> lupasOff = new List<int>() { 0 };
                List<string> editValue = new List<string>() { variables[variables.Count-1] };
                Thread.Sleep(1000);
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editValue, lupasOff);
                Thread.Sleep(1000);
            }
        }
    }
}
