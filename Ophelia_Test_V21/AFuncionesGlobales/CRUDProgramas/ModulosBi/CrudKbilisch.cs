using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbilisch
    {

        public static void KBiLischCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> val, int flag, int fieldIndex, string editValue)
        {

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            int i = 0;


            if (flag == 0)
            {
                foreach (var elem in editFields)
                {

                    if (elem.Enabled)
                    {
                        elem.ClearCache();
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        desktopSession.Keyboard.SendKeys(val[i]);
                        i++;
                    }
                    Thread.Sleep(1000);
                }
            }
            else
            {
                editFields[fieldIndex].Clear();
                desktopSession.Mouse.MouseMove(editFields[fieldIndex].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[fieldIndex].SendKeys(editValue);
            }


        }

        public static void KBiLischCRUDDetalle(WindowsDriver<WindowsElement> desktopSession, List<string> valores, List<int> numerotabs) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            //System.Diagnostics.//Debugger.Launch();




            desktopSession.Mouse.MouseMove(editFields[1].Coordinates, editFields[1].Size.Width - 10, editFields[1].Size.Height / 2);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            for (int i = 0; i < valores.Count - 1; i++)
            {
                desktopSession.Keyboard.SendKeys(valores[i]);
                if (i == 0) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                if (i < numerotabs.Count)
                {
                    for (int tbs = 0; tbs < numerotabs[i]; tbs++)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                    }
                }

            }
        }


        public static void KBiLischCRUDDetalleEdit(WindowsDriver<WindowsElement> desktopSession, int numneroTab, string editValue) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            //System.Diagnostics.//Debugger.Launch();

            desktopSession.Mouse.MouseMove(editFields[1].Coordinates, editFields[1].Size.Width/2, editFields[1].Size.Height / 2);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            for (int i = 0; i < numneroTab; i++) {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Tab + OpenQA.Selenium.Keys.Shift);
                Thread.Sleep(1000);
            }

            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
            desktopSession.Keyboard.SendKeys(editValue);
            Thread.Sleep(1000);



        }

    }
}
