using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbidocum
    {
        /// <summary>
        /// Flag = 0 para agregar
        /// Flag = 1 para editar
        /// </summary>
        /// <param name="desktopSession"></param>
        /// <param name="datosCabecera"></param>
        /// <param name="flag"></param>
        /// <param name="fieldIndex"></param>
        /// <param name="editValue"></param>
        public static void KBiDocumCRUDCabecera(WindowsDriver<WindowsElement> desktopSession, List<string> datosCabecera, int flag, int fieldIndex, string editValue)
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
                
                for (int index = fieldIndex; index < editFields.Count; index++) {
                    desktopSession.Mouse.MouseMove(editFields[index].Coordinates);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(datosCabecera[i]);
                    i++;
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


        public static void KBiDocumCRUDDetalle(WindowsDriver<WindowsElement> desktopSession, string datosDetalle) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBGrid");

            desktopSession.Mouse.MouseMove(editFields[0].Coordinates, 20, 30);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            desktopSession.Keyboard.SendKeys(datosDetalle);
            //System.Diagnostics.//Debugger.Launch();
            Thread.Sleep(3000);
            desktopSession.Mouse.MouseMove(editFields[0].Coordinates, 200, 30);
            desktopSession.Mouse.Click(null);

        }


        public static void changeWindow(WindowsDriver<WindowsElement> desktopSession, string name) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByName(name);

            desktopSession.Mouse.MouseMove(editFields[0].Coordinates, editFields[0].Size.Width/2, editFields[0].Size.Height/2);
            Thread.Sleep(3000);
            desktopSession.Mouse.Click(null);
        }


    }
}
