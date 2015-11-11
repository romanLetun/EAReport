using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EAConnect
{
    class Program
    {
        public static LinkedList<EA.Package> packages = new LinkedList<EA.Package>();
        public static EA.Repository r = new EA.Repository();

        static void Main(string[] args)
        {                      
            int packageID;
            String[] list;
            
            EA.DocumentGenerator doc = r.CreateDocumentGenerator();

            try
            {
                Console.WriteLine("Введите наименование файла модели");

                r.OpenFile(Environment.CurrentDirectory + "\\" + Console.ReadLine());
                EA.Project project = new EA.Project();

                Console.WriteLine("Введите наименование выходного файла");
                String path = Environment.CurrentDirectory + "\\" + Console.ReadLine();

                Console.WriteLine("Введите путь к пакету. разделительный символ - \\");
                list = Console.ReadLine().Split('\\');
                
                packageID = getPackageId(list);

                Console.WriteLine("Документирую!");
                ////////////////

                project.RunReport(r.GetPackageByID(packageID).PackageGUID, "sequential", path);                

                Console.WriteLine("OK");
            }
            catch (Exception e)
            {              
                Console.WriteLine(e.GetBaseException());
                Console.ReadLine();              
            }
            finally
            {
                r.CloseFile();
                r.Exit();
            }
        }

        public static int getPackageId(String[] list)
        {
            EA.Package package = r.GetPackageByID(1);

            for (short i = 1; i < list.Length; i++)
            {
                for (short j = 0; j < package.Packages.Count; j++)
                {
                    if (package.Packages.GetAt(j).Name == list[i])
                    {
                        package = package.Packages.GetAt(j);
                        break;
                    }
                }
            }

            return package.PackageID;
        }
    }
}
