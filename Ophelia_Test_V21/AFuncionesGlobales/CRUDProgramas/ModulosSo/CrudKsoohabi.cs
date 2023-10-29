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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoohabi
    {
        public static void KSoOhabiCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                editFields[2].Click();
                editFields[2].Clear();
                Thread.Sleep(1000);
                editFields[2].SendKeys(variables[0]);
                Thread.Sleep(1000);
                memos[0].Click();
                memos[0].Clear();
                Thread.Sleep(1000);
                memos[0].SendKeys(variables[1]);
                Thread.Sleep(1000);

            }
            else {
                Thread.Sleep(1000);
                memos[0].Click();
                memos[0].Clear();
                Thread.Sleep(1000);
                memos[0].SendKeys(variables[variables.Count-1]);
            }
        }
    }
}
