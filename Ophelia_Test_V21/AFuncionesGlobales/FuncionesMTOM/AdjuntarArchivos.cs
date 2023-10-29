
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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using System.Drawing.Imaging;
using System.IO;
using Ophelia_Test_V21.AFuncionesGlobales.FuncionesMTOM;

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesMTOM
{
    class AdjuntarArchivos:FuncionesVitales
    {

        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static void flechaRoja(WindowsDriver<WindowsElement> desktopSession)
        {
            List<Point> coordenadas = coordinatesFinder2(desktopSession);
            int coordx = coordenadas[0].X;
            int coordy = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordx, coordy);
            desktopSession.Mouse.Click(null);
        }
        
        public static void controlVentana(WindowsDriver<WindowsElement> desktopSession, string bandera, string CodEmpleado, string Database, string User, string Table, string CodEmpresa, string Almacenamiento, string RutaCP, string RutaFinalFTP, List<string> ValidaSubida, List<string> ErrorValidacion, List<string> ValidaDescarga, string file, string BanderaSmpeople, string URL, string PassSelf, string url2)
        {
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmAdjuntar");
            var BotonList = rootSession.FindElementsByClassName("TBitBtn");
            List<string> tiposDoc = new List<string>() { "Excel_MTOM", "PDF_MTOM", "Word_MTOM" };

            if(bandera=="1")
            {
                foreach (var btn in BotonList)
                {
                    if (btn.Text == "Adicionar")
                    {
                        Thread.Sleep(2000);
                        for (int i = 0; i <= 2; i++)
                        {
                            btn.Click();
                            Thread.Sleep(1500);
                            WindowsDriver<WindowsElement> rootSession2 = RootSession();
                            rootSession2.Keyboard.SendKeys(tiposDoc[i]);
                            Thread.Sleep(500);
                            rootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                        }
                        ////VALIDA SUBIDA DE ADJUNTOS//////
                        ValidaSubida = ValidaAdjuntos(CodEmpleado, Database, User, Table, CodEmpresa, tiposDoc, Almacenamiento, RutaCP, RutaFinalFTP);
                        ValidaSubida.ForEach(u => ErrorValidacion.Add(u));
                    }
                }
            }
            else if (bandera == "2" || bandera == "3")
            {
                var gridInterna = rootSession.FindElementsByClassName("TDBGrid");
                new Actions(rootSession).MoveToElement(gridInterna[0], 0, 0).MoveByOffset(50, 20).ClickAndHold().Click().Perform();
                Thread.Sleep(1000);
                new Actions(rootSession).Click().Perform();

                string excel = @"Reporte_" + "MTOM" + "_" + Hora() + ".pdf";
                foreach (var btn in BotonList)
                {
                    if (btn.Text == "Descargar" || btn.Text == "Descargar Todo")
                    {
                        if (btn.Text == "Descargar")
                        {
                            btn.Click();
                            WindowsDriver<WindowsElement> rootSession2 = RootSession();
                            rootSession2.Keyboard.SendKeys(excel);
                            Thread.Sleep(1000);
                            rootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            rootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        }
                        else
                        {
                            btn.Click();
                            //WindowsDriver<WindowsElement> rootSession2 = RootSession();
                            //var ventana2 = rootSession.FindElementsByName("Buscar carpeta");
                            var ventana2 = rootSession.FindElementsByName("Browse for Folder");
                            ////Debugger.Launch();
                            ventana2[0].Click();
                            ventana2[0].SendKeys("Disco");
                            Thread.Sleep(500);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            for (int i = 1; i <= 4; i++)
                            {
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(2000);
                            }
                            //C:\Users\kauto1\Downloads\Word_MTOM.docx
                            //C:\Users\kauto1\Downloads\Excel_MTOM.xlsx
                            //C:\Users\kauto1\Downloads\PDF_MTOM.pdf
                            List<string> tiposDocRut = new List<string>() { @"C:\Users\kauto1\Downloads\Excel_MTOM.xlsx", @"C:\Users\kauto1\Downloads\PDF_MTOM.pdf", @"C:\Users\kauto1\Downloads\Word_MTOM.docx" };
                            ////VALIDA DESCARGAR TODO//////
                            ValidaDescarga = AdjuntarArchivos.ValidaDescarga(tiposDocRut, file);
                            ValidaDescarga.ForEach(u => ErrorValidacion.Add(u));
                        }
                    }
            
                }
            }
            else if (bandera == "4")
            {
                foreach (var btn in BotonList)
                {
                    if (btn.Text == "Abrir")
                    {
                        btn.Click();
                        Thread.Sleep(5000);
                        LimpiarProcesos();
                    }
                }
                //////////Aqui llamar a la clase SmartPeopleMTOM y sus metodos/////////////////
                if (BanderaSmpeople == "1")
                {
                    errors = SmartPeopleMTOM.irUrlSmartPeople(URL, CodEmpleado, PassSelf, url2, tiposDoc, file);
                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                }
                /////Enviar una bandera y preguntar si se llama o no al selfservice
                ///if(bandera) true o false
                ///{
                ///   SmartPeopleMTOM.irUrlSmartPeople(URL, CodEmpleado, PassSelf, url2, tiposDoc, file);
                ///}
                //SmartPeopleMTOM.irUrlSmartPeople(URL, CodEmpleado, PassSelf, url2, tiposDoc, file);
                
            }
            else if (bandera == "5")
            {
                foreach (var btn in BotonList)
                {
                    if (btn.Text == "Eliminar")
                    {
                        btn.Click();
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    }
                }
            }
            else if (bandera == "6")
            {
                foreach (var btn in BotonList)
                {
                    if (btn.Text == "Eliminar Todo")
                    {
                        btn.Click();
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    }
                }
            }
            else
            {
                var close = rootSession.FindElementsByName("Cerrar");
                close[0].Click();
            }
        }

        private static List<Point> coordinatesFinder2(WindowsDriver<WindowsElement> session)
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

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);

            for (int x = 0; x < bmpSource.Width; x++)
            {
                for (int y = 0; y < bmpSource.Height; y++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);

                    if (clrPixel.R == 237 && clrPixel.G == 28 && clrPixel.B == 36)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 20 || y > lasty + 20)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                        }
                    }
                }
            }
            return coordenadas;
        }

        public static List<string> ValidaDescarga(List<string> Documents, string file)
        {
            List<string> Errores = new List<string>();
            foreach (string path in Documents)
            {
                if (!File.Exists(path))
                {
                    Errores.Add(string.Format("Ocurrió un error al descargar el archivo adjunto. File: {0}", path));
                }
                else
                {
                    Process.Start(path);
                    Thread.Sleep(6000);
                    //Screenshot("Archivo Descargado MTOM", true, file);
                    Thread.Sleep(1000);
                    LimpiarProcesos();
                    //Thread.Sleep(1000);
                    //File.Delete(path);
                }
            }
            return Errores;
        }



    }
}
