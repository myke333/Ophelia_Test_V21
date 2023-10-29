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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiroles : FuncionesVitales
    {
        public static void CRUDKBiRoles(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");


            List<string> data = new List<string>() { crudVars[0] };

            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    desktopSession.Mouse.MouseMove(ElementList[i].Coordinates, 10, 10);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(1000);
            //}
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                Thread.Sleep(1000);
                ElementList[11].SendKeys(crudVars[1]);
                ElementList[5].SendKeys(crudVars[2]);
                ElementList[7].SendKeys(crudVars[3]);
                ElementList2[0].SendKeys(crudVars[5]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(crudVars[4]);
            }
        }

        public static void CRUDDetalle1KBiRoles(WindowsDriver<WindowsElement> desktopSession, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBMemo");
            var Element2 = desktopSession.FindElementsByClassName("TTabSheet");
            var Element3 = Element2[0].FindElementsByClassName("TKNavegador");
            //EDITAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
            desktopSession.Mouse.Click(null);

            ElementList[0].SendKeys(crudVarsdet1[0]);

            //DESCARTAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 60, 15);
            desktopSession.Mouse.Click(null);

            //EDITAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
            desktopSession.Mouse.Click(null);

            ElementList[0].SendKeys(crudVarsdet1[0]);

            //ACEPTAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 30, 15);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);

                    
        }

        public static void CRUDDetalle2KBiRoles(WindowsDriver<WindowsElement> desktopSession, List<string> crudVarsdet2, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");

            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    desktopSession.Mouse.MouseMove(ElementList[i].Coordinates, 10, 10);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(1000);
            //}
            WindowsDriver<WindowsElement> rootSession = null;
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 40, 10);
            desktopSession.Mouse.DoubleClick(null);
            //desktopSession.Mouse.Click(null);

            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmFuncion");

            var ed = rootSession.FindElementsByClassName("TDBMemo");
            var Element2 = rootSession.FindElementsByClassName("TKNavegador");
            var Element3 = rootSession.FindElementsByName("Cerrar");
            Thread.Sleep(500);
            
            ed[0].SendKeys(crudVarsdet2[0]);
            Thread.Sleep(500);

            //ACEPTAR
            rootSession.Mouse.MouseMove(Element2[0].Coordinates, 55, 15);
            rootSession.Mouse.Click(null);
            //EDITAR
            rootSession.Mouse.MouseMove(Element2[0].Coordinates, 20, 15);
            rootSession.Mouse.Click(null);

            ed[0].Clear();
            ed[0].SendKeys(crudVarsdet2[1]);
            //DESCARTAR
            rootSession.Mouse.MouseMove(Element2[0].Coordinates, 95, 15);
            rootSession.Mouse.Click(null);
            //EDITAR
            rootSession.Mouse.MouseMove(Element2[0].Coordinates, 20, 15);
            rootSession.Mouse.Click(null);

            ed[0].Clear();
            ed[0].SendKeys(crudVarsdet2[1]);
            //ACEPTAR
            rootSession.Mouse.MouseMove(Element2[0].Coordinates, 55, 15);
            rootSession.Mouse.Click(null);
            //CERRAR
            rootSession.Mouse.MouseMove(Element3[0].Coordinates, 20, 15);
            rootSession.Mouse.Click(null);


        }

        public static void CRUDDetalle3KBiRoles(WindowsDriver<WindowsElement> desktopSession, List<string> crudVarsdet3, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBMemo");
            var Element2 = desktopSession.FindElementsByClassName("TTabSheet");
            var Element3 = Element2[0].FindElementsByClassName("TKNavegador");

            //EDITAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
            desktopSession.Mouse.Click(null);

            ElementList[0].SendKeys(crudVarsdet3[0]);

            //DESCARTAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 60, 15);
            desktopSession.Mouse.Click(null);

            //EDITAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 5, 15);
            desktopSession.Mouse.Click(null);

            ElementList[0].SendKeys(crudVarsdet3[0]);

            //ACEPTAR
            desktopSession.Mouse.MouseMove(Element3[0].Coordinates, 30, 15);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            
        }
        static public void ventanaok(WindowsDriver<WindowsElement> desktopSession)
        {







            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");







            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 986, 361);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



        }
        public static void CRUDDetalle4KBiRoles(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet4, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 104, 33);
                desktopSession.Mouse.DoubleClick(null);

                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");

                var ed = rootSession.FindElementsByClassName("TEdit");
                ed[0].SendKeys(crudVarsdet4[0]);
                Thread.Sleep(500);
                ed[1].SendKeys(crudVarsdet4[0]);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 104, 33);
                desktopSession.Mouse.Click(null);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var aceptar = rootSession.FindElementsByName("Aceptar");
                rootSession.Mouse.MouseMove(aceptar[0].Coordinates, 30, 10);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
               var ElementListd = desktopSession.FindElementsByClassName("TDBGrid");
                desktopSession.Mouse.MouseMove(ElementListd[0].Coordinates,490, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);

                Thread.Sleep(1000);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                aceptar = rootSession.FindElementsByName("Aceptar");
                rootSession.Mouse.MouseMove(aceptar[0].Coordinates, 30, 10);
                rootSession.Mouse.Click(null);
            }
            else
            {
                
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                //var allFrame = rootSession.FindElementsByClassName("IEFrame");
                //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();
                rootSession = PruebaCRUD.RootSession();
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 104, 33);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Mouse.DoubleClick(null);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");

                var ed = rootSession.FindElementsByClassName("TEdit");
                ed[0].SendKeys(crudVarsdet4[1]);
                Thread.Sleep(500);
                ed[1].SendKeys(crudVarsdet4[1]);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);


            }
        }
        public static bool coordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B, int fil, int col, int tipo)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = FuncionesVitales.Hora();
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            bool val = false;

            //Color clrPixel = bmpSource.GetPixel(fil, col);

            if (tipo == 2)
            {
                Color clrPixel = bmpSource.GetPixel(fil, col);
                if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                {
                    val = true;
                }
            }
            else if (tipo == 1)
            {
                int cont = 0;
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    for (int y = 0; y < bmpSource.Height; y++)
                    {
                        Color clrPixel = bmpSource.GetPixel(x, y);
                        if (cont == 0)
                        {
                            if (clrPixel.R == 203 && clrPixel.G == 219 && clrPixel.B == 234)
                            {
                                cont = 1;
                                val = false;
                                break;
                            }
                        }
                    }
                    if (cont != 0) { break; }
                }
                if (cont == 0) { val = true; }
            }

            return val;
        }

        public static void CRUDDetalle5KBiRoles(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet5, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
            desktopSession.Mouse.Click(null);
            int cont = 0;
            if (tipo == 1)
            {
                desktopSession.Keyboard.SendKeys(crudVarsdet5[cont]);
                for (int i = 0; i < 13; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudVarsdet5[cont]);
                    cont += 1;
                    Thread.Sleep(500);             
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                for (int i = 0; i < 13; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Left);

                }
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet5[14]);
            }
        }

        public static void CRUDDetalle6KBiRoles(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet6, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
            desktopSession.Mouse.Click(null);
            if (tipo == 1)
            {
                desktopSession.Keyboard.SendKeys(crudVarsdet6[0]);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(crudVarsdet6[1]);
            }
        }

        public static void CRUDDetalle7KBiRoles(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet7, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 30);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmDesComun");
            var ElementList2 = rootSession.FindElementsByClassName("TDBMemo");            
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet7[0]);
            }
            else
            {
                ElementList2[0].SendKeys(crudVarsdet7[1]);
            }
            rootSession.Mouse.MouseMove(ElementList2[0].Coordinates, 50, 30);
            rootSession.Mouse.Click(null);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }

    }
}
