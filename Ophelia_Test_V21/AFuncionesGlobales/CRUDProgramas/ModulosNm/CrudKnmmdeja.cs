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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmmdeja
    {
        public static void KNmMdejaCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag, string file) {
            
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dropFields = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            //TDBCmbCodigoBox


            if (flag == 0)
            {
                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(editFields[4].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[4].Clear();
                desktopSession.Keyboard.SendKeys(variables[2]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(editFields[5].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[5].Clear();
                desktopSession.Keyboard.SendKeys(variables[1]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(dropFields[0].Coordinates);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
            }
            else {
                List<string> editLupa = new List<string>() { variables[variables.Count - 1] };
                PruebaCRUD.LupaDinamica(desktopSession, editLupa);
                Thread.Sleep(2000);
            }
        }
    }
}
