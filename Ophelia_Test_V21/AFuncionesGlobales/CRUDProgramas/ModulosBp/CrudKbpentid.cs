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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbpentid
    {
        public static void KBpEntidCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> memoFields = desktopSession.FindElementsByClassName("TDBMemo");
            List<int> fieldIndex = new List<int>() { 9, 8, 7, 6, 5, 17, 16, 15, 11, 3, 2, 4 };

            if (flag == 0)
            {


                for (int i = 0; i < fieldIndex.Count; i++)
                {
                    editFields[fieldIndex[i]].Click();
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    editFields[fieldIndex[i]].SendKeys(variables[i]);
                    Thread.Sleep(1000);
                }

                memoFields[0].Click();
                memoFields[0].Clear();
                Thread.Sleep(1000);
                memoFields[0].SendKeys(variables[variables.Count - 2]);
                Thread.Sleep(1000);

            }
            else {
                Thread.Sleep(1000);
                memoFields[0].Click();
                memoFields[0].Clear();
                Thread.Sleep(1000);
                memoFields[0].SendKeys(variables[variables.Count - 1]);
                Thread.Sleep(1000);
            }

        }
    }
}
