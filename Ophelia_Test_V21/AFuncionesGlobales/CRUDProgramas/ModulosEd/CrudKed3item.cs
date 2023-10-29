using OpenQA.Selenium.Appium.Windows;
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
    class CrudKed3item
    {
        public static void KEd3itemCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> memos = desktopSession.FindElementsByClassName("TDBMemo");
            if (flag == 0)
            {
                editFields[8].Click();
                editFields[8].Clear();
                Thread.Sleep(1000);
                editFields[8].SendKeys(variables[3]);
                Thread.Sleep(1000);
                memos[0].Click();
                memos[0].Clear();
                Thread.Sleep(1000);
                memos[0].SendKeys(variables[4]);
                Thread.Sleep(1000);
                PruebaCRUD.LupaDinamica(desktopSession, variables);

            }
            else {
                Thread.Sleep(1000);
                memos[0].Click();
                memos[0].Clear();
                Thread.Sleep(1000);
                memos[0].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }
        }
    }
}
