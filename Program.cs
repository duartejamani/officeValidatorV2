using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using System.Diagnostics;
using System.Security.Principal;
using System.Reflection;

using System.IO;
using System.Threading;

namespace officeValidatorV2
{
    class Program
    {
        public static int version = 0;
        const char Comillas = '"';
        const string ruta = @"..\root\Licenses16\%x";
        public static string path64Bits = @"C:\Program Files\Microsoft Office\Office16";
        public static string path32Bits = @"C:\Program Files (x86)\Microsoft Office\Office16";
        public static bool rutabool = false;

        /// <summary>
        /// Nuevos metodos 
        /// </summary>
        /// 
        ///
        public static bool validateVersion(string path)
        {
            if (Directory.Exists(path))
            {
                rutabool = true;
                return true;
            }
            else
            {
                rutabool = false;
                return false;
            }
        }
        public static void status(string Path)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.WorkingDirectory = Path;
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();

            cmd.StandardInput.WriteLine("cscript ospp.vbs /dstatus");
            cmd.StandardInput.Flush();

            cmd.StandardInput.Close();
            cmd.WaitForExit();
            Console.WriteLine(cmd.StandardOutput.ReadToEnd());
        }

        static void ExecuteCommand(string _Command, string path)
        {
            //Indicamos que deseamos inicializar el proceso cmd.exe junto a un comando de arranque. 
            //(/C, le indicamos al proceso cmd que deseamos que cuando termine la tarea asignada se cierre el proceso).
            //Para mas informacion consulte la ayuda de la consola con cmd.exe /? 
            System.Diagnostics.ProcessStartInfo procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/c " + _Command);
            procStartInfo.WorkingDirectory = path;
            // Indicamos que la salida del proceso se redireccione en un Stream
            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            //Indica que el proceso no despliegue una pantalla negra (El proceso se ejecuta en background)
            procStartInfo.CreateNoWindow = false;
            //Inicializa el proceso
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo = procStartInfo;
            proc.StartInfo.WorkingDirectory = path;
            proc.Start();
            //Consigue la salida de la Consola(Stream) y devuelve una cadena de texto
            string result = proc.StandardOutput.ReadToEnd();
            //Muestra en pantalla la salida del Comando
            Console.WriteLine(result);
        }
        public static void callValidate(string path)
        {
            Console.WriteLine("Convirtiendo en versión RETAIL a VL (Licencia por Volumen) ...");
            string comando0 = "for /f %x in ('dir /b ..\\root\\Licenses16\\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:" + Comillas + "..\\root\\Licenses16\\%x" + Comillas;
            //string comando0 = @"for / f % x in ('dir /b ..\\root\\Licenses16\\proplusvl_kms*.xrm-ms') do cscript ospp.vbs / inslic:"""..\\root\\Licenses16\\% x"""\ ;
            ExecuteCommand(comando0, path);

            Thread.Sleep(3000);
            Console.WriteLine("Instalando clave de producto ...");
            string comando1 = "cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
            ExecuteCommand(comando1, path);
            
            Thread.Sleep(3000);
            Console.WriteLine("Desinstalando clave parcial de producto ...");
            string comando2 = "cscript ospp.vbs /unpkey:BTDRB >nul";
            ExecuteCommand(comando2, path);

            Thread.Sleep(3000);
            Console.WriteLine("Desinstalando clave parcial de producto ...");
            string comando3 = "cscript ospp.vbs /unpkey:KHGM9 >nul";
            ExecuteCommand(comando3, path);

            Thread.Sleep(3000);
            Console.WriteLine("Desinstalando clave parcial de producto ...");
            string comando4 = "cscript ospp.vbs /unpkey:CPQVG >nul";
            ExecuteCommand(comando4, path);

            Thread.Sleep(3000);
            Console.WriteLine("Estableciendo host KMS  ...");
            string comando5 = "cscript ospp.vbs /sethst:kms8.msguides.com";
            ExecuteCommand(comando5, path);

            Thread.Sleep(3000);
            Console.WriteLine("Estableciendo puerto KMS  ...");
            string comando6 = "cscript ospp.vbs /setprt:1688";
            ExecuteCommand(comando6, path);

            Thread.Sleep(3000);
            Console.WriteLine("Activando ...");
            string comando7 = "cscript ospp.vbs /act"; 
            ExecuteCommand(comando7, path);
        }

        //public static void validateOffice(string path)
        //{
        //    Process cmd = new Process();
        //    cmd.StartInfo.FileName = "cmd.exe";
        //    cmd.StartInfo.WorkingDirectory = @path;
        //    cmd.StartInfo.RedirectStandardInput = true;
        //    cmd.StartInfo.RedirectStandardOutput = true;
        //    cmd.StartInfo.CreateNoWindow = true;
        //    cmd.StartInfo.UseShellExecute = false;
        //    cmd.Start();

        //    //cmd.StandardInput.WriteLine(@"for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:" + Comillas + @"..\root\Licenses16\%x" + Comillas);
        //    //cmd.StandardInput.WriteLine(@"for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:" + Comillas + ruta + Comillas);
        //    cmd.StandardInput.WriteLine("for /f %x in ('dir /b ..\\root\\Licenses16\\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:" + Comillas + "..\\root\\Licenses16\\%x" + Comillas);
        //    cmd.StandardInput.Flush();
        //    //Console.WriteLine("1");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99");
        //    //cmd.StandardInput.Flush();
        //    ////Console.WriteLine("2");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /unpkey:BTDRB >nul");
        //    //cmd.StandardInput.Flush();
        //    ////Console.WriteLine("3");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /unpkey:KHGM9 >nul");
        //    //cmd.StandardInput.Flush();
        //    ////Console.WriteLine("4");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /unpkey:CPQVG >nul");
        //    //cmd.StandardInput.Flush();
        //    ////Console.WriteLine("5");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /sethst:kms8.msguides.com");
        //    //cmd.StandardInput.Flush();
        //    ////Console.WriteLine("6");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /setprt:1688");
        //    //cmd.StandardInput.Flush();
        //    ////Console.WriteLine("7");
        //    //cmd.StandardInput.WriteLine("cscript ospp.vbs /act");
        //    //cmd.StandardInput.Flush();
        //    //Console.WriteLine("8");


        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /unpkey:BTDRB >nul");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /unpkey:KHGM9 >nul");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /unpkey:CPQVG >nul");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /sethst:kms8.msguides.com");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /setprt:1688");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.WriteLine("cscript ospp.vbs /act");
        //    cmd.StandardInput.Flush();
        //    cmd.StandardInput.Close();
        //    cmd.WaitForExit();
            
        //    Console.WriteLine(cmd.StandardOutput.ReadToEnd());
        //    return;

        //}


        private static bool MainMenu()
        {
            //Validar validar = new Validar();
            //Console.Clear();
            Console.WriteLine("1 .- Verificar Status");
            Console.WriteLine("2 .- Validar office 32 & 64 bits");
            Console.WriteLine("3 .- Exit");
            //Console.WriteLine("4 .- Validar office 32 & 64 bits");
            //Console.WriteLine(path32Bits);
            //Console.WriteLine(path64Bits);
            //Console.WriteLine("for /f %x in ('dir /b ..\\root\\Licenses16\\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:" + Comillas + "..\\root\\Licenses16\\%x" + Comillas);
            Console.Write("\r\nSelecciona una opción : ");

            switch (Console.ReadLine())
            {
                case "1":
                    Console.WriteLine("///////////////////////////////////////////////////////////////// \n ");
                    Console.WriteLine("selecciono 1 ");
                    Console.WriteLine("Verificando validación  ... ...");
                    //verificarStatus();

                    if (validateVersion(path32Bits))
                    {
                        Console.WriteLine("es de 32 bits");
                        status(path32Bits);
                    }
                    else if (validateVersion(path64Bits))
                    {
                        Console.WriteLine("es 64 bits");
                        status(path64Bits);
                    }
                    else
                    {
                        Console.WriteLine("No cuenta con una versión de office instalada en este ordenador .....");
                    }
                    Console.WriteLine(" \n ");
                    Console.WriteLine("///////////////////////////////////////////////////////////////// \n ");
                    return true;
                case "2":
                    Console.WriteLine("///////////////////////////////////////////////////////////////// \n ");
                    Console.WriteLine("selecciono 2 ");
                    Console.WriteLine("Validando versión 32 & 64 bits ... ...");
                    if (validateVersion(path32Bits))
                    {
                        Console.WriteLine("es de 32 bits");
                        Console.WriteLine(path32Bits);
                        callValidate(path32Bits);
                        //validar.validateOffice(path32Bits);
                    }
                    else if (validateVersion(path64Bits))
                    {
                        Console.WriteLine("es 64 bits");
                        Console.WriteLine(path64Bits);
                        callValidate(path64Bits);
                        //validateOffice(path64Bits);
                        //validar.validateOffice(path64Bits);
                    }
                    else
                    {
                        Console.WriteLine("No cuenta con una versión de office instalada en este ordenador .....");
                    }
                    Console.WriteLine(" \n ");
                    Console.WriteLine("///////////////////////////////////////////////////////////////// \n ");
                    return true;
                case "3":
                    System.Environment.Exit(1);
                    return false;
                default:
                    return true;
            }
        }

        private static bool IsRunAsAdmin()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(id);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void Main(string[] args)
        {


            if (!IsRunAsAdmin())
            {
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.UseShellExecute = true;
                proc.WorkingDirectory = Environment.CurrentDirectory;
                proc.FileName = Assembly.GetEntryAssembly().CodeBase;

                foreach (string arg in args)
                {
                    proc.Arguments += String.Format("\"{0}\" ", arg);
                }

                proc.Verb = "runas";

                try
                {
                    Process.Start(proc);
                }
                catch
                {
                    Console.WriteLine("This application requires elevated credentials in order to operate correctly!");
                }
            }
            else
            {
                //Console.WriteLine("Hello World!");
                //Process cmd = new Process();
                //cmd.StartInfo.FileName = "cmd.exe";
                //cmd.StartInfo.WorkingDirectory = @"C:";
                //cmd.StartInfo.RedirectStandardInput = true;
                //cmd.StartInfo.RedirectStandardOutput = true;
                //cmd.StartInfo.CreateNoWindow = true;
                //cmd.StartInfo.UseShellExecute = false;        
                //cmd.Start();

                //cmd.StandardInput.WriteLine("dir");
                //cmd.StandardInput.Flush();

                //cmd.StandardInput.Close();
                //cmd.WaitForExit();
                //Console.WriteLine(cmd.StandardOutput.ReadToEnd());
                bool showMenu = true;
                while (showMenu)
                {
                    showMenu = MainMenu();
                }
                //Normal program logic...
            }




        }

    }
}
