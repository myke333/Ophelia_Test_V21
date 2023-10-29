using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmcvcon
    {
        public static void CrudKNmcvcon(WindowsDriver<WindowsElement> desktopSession, int Tipo, string numDiasTemp, string email, string numDias, string asunto, string contenido, string Editar)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementByClassName("TDBMemo");

            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Terminación de Contrato").Coordinates);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Jefe Inmediato").Coordinates);
                desktopSession.Mouse.Click(null);

                //ElementList[2].SendKeys(numDiasTemp);
                
                ElementList[2].SendKeys(email);

                //ElementList[4].SendKeys(numDias);
                ElementList[4].SendKeys(asunto);
                ElementList1.SendKeys(contenido);
            }
            else if (Tipo == 2)
            {
                ElementList1.Clear();
                ElementList1.SendKeys(Editar);
            }
        }
    }
}
