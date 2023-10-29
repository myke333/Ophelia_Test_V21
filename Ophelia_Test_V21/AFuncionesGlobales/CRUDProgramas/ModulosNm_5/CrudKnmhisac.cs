using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmhisac : FuncionesVitales
    {
        public static void CRUDKNmHisac(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file, int tipo)
        {
            var Element = desktopSession.FindElementsByClassName("TEdit");
            var Element2 = desktopSession.FindElementsByClassName("TScrollBox");
            var Element3 = desktopSession.FindElementsByName("Trasladar");
            var Element4 = desktopSession.FindElementsByName("Acumulados");
            WindowsDriver<WindowsElement> rootSession = null;

            if (tipo==1)
            {
                Thread.Sleep(500);
                Element[1].SendKeys(crudVars[0]);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Digitar Fecha Inicial", true, file);

                //Click Datos
                desktopSession.Mouse.MouseMove(Element2[0].Coordinates, 590, 70);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
            if (tipo == 2)
            {
                //Click trasladar
                Element3[0].Click();
                //Si trasladar tiene ventana QBE
                try
                {
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
                    Thread.Sleep(500);
                    var Element5 = rootSession.FindElementsByClassName("TInplaceEdit");
                    Element5[0].SendKeys(crudVars[2]);
                    var Element6 = rootSession.FindElementsByClassName("TPanel");
                    rootSession.Mouse.MouseMove(Element6[5].Coordinates, 40, 10);
                    rootSession.Mouse.Click(null);
                }
                catch 
                {
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("Trasladar datos", true, file);
                    Thread.Sleep(500);
                    try
                    {
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                        var Element7 = rootSession.FindElementsByName("Yes");
                        Thread.Sleep(500);
                        newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
                        Thread.Sleep(500);
                        Element7[0].Click();
                    }
                    catch 
                    {
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                        Thread.Sleep(500);
                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 25, (allFrame[0].Size.Height / 2) + 35).DoubleClick().Perform();
                        Thread.Sleep(500);
                        newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
                        Thread.Sleep(500);
                        try
                        {
                            rootSession = PruebaCRUD.RootSession();
                            rootSession = PruebaCRUD.ReloadSession(rootSession, "TMessageForm");
                            var Element7 = rootSession.FindElementsByName("OK");
                            Element7[0].Click();
                        }
                        catch
                        {
                            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2), (allFrame[0].Size.Height / 2) + 30).DoubleClick().Perform();
                        }
                        
                    }
                }
            }
            if (tipo == 3)
            {
                //Click Parametros
                desktopSession.Mouse.MouseMove(Element2[0].Coordinates, 500, 70);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                Element4[0].Click();
                Thread.Sleep(500);
                //Click Datos
                desktopSession.Mouse.MouseMove(Element2[0].Coordinates, 590, 70);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }              
        }
    }
}
