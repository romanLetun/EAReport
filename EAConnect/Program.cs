//EAReport 1.0 Letunovskii Roman(c)
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace EAConnect
{
    class Program
    {
        const char delimiter = '~';

        static void Main(string[] args)
        {    
            int packageID;
            String[] list;
            String template;
            String path = "test.docx";
            String docTemplatePath = "Le.dotx";
            TextWriter errWriter = Console.Error;
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            Word.Application app = new Word.Application();
            //Стили
            Word.Style pictureName = null;
            Word.Style tableName = null;
            Word.Style header1 = null;
            Word.Style header2 = null;
            Word.Style header3 = null;
            Word.Style header4 = null;
            Word.Style header5 = null;

            try
            {
                if (args.Length < 1)
                {
                    throw new ReportException("Запуск без параметров", 1);
                }

                if (args[0] == "h" || args[0] == "?")
                {
                    printHelp();
                    Environment.Exit(1);
                }

                else
                    if (args.Length == 5)
                    {
                        EA.Repository r = new EA.Repository();
                        EA.Project project = new EA.Project();
                        //Открываем файл модели
                        if (File.Exists(Environment.CurrentDirectory + "\\" + args[0]))
                            r.OpenFile(Environment.CurrentDirectory + "\\" + args[0]);
                        else
                            throw new ReportException("Файл модели не найден", 1);

                        //Выходной файл
                        if (args[1].Substring(args[1].Length - 5) == ".docx")
                            path = Environment.CurrentDirectory + "\\" + args[1];
                        else
                            throw new ReportException("Наименование выходного файла должно соответствовать шаблону \\w+\\.docx", 1);

                        //Путь к документируемому пакету
                        list = args[2].Split(delimiter);

                        if (list.Length < 2)
                            throw new ReportException("Путь к документируемому пакету должен содержать хотя бы один символ \"~\"", 1);

                        //Наименование шаблона документации
                        template = args[3];

                        //Получаем ID пакета
                        packageID = getPackageId(list, r);

                        //Запуск документации по шаблону
                        project.RunReport(r.GetPackageByID(packageID).PackageGUID, template, path);

                        r.CloseFile();
                        r.Exit();                        
                    }
                    else
                        throw new ReportException("Не верное количество параметров", 1);

                docTemplatePath = args[4];

                if (File.Exists(path) && File.Exists(docTemplatePath))
                {
                    Word.Document doc = app.Documents.Open(path);
                    Word.Document docTemplate = app.Documents.Open(docTemplatePath);

                    foreach (Word.Style s in docTemplate.Styles)
                    {
                        switch (s.NameLocal)
                        {
                            case "lePictureName":
                                pictureName = s;
                                break;
                            case "leTableName":
                                tableName = s;
                                break;
                            case "leHeader1":
                                header1 = s;
                                break;
                            case "leHeader2":
                                header2 = s;
                                break;
                            case "leHeader3":
                                header3 = s;
                                break;
                            case "leHeader4":
                                header4 = s;
                                break;
                            case "leHeader5":
                                header5 = s;
                                break;
                        }
                    }

                    foreach (Word.Paragraph p in doc.Paragraphs)
                    {
                        switch ((String)p.get_Style().NameLocal)
                        {
                            case "leCode":
                                p.set_Style(pictureName);
                                break;
                            case "leUsual":
                                p.set_Style(tableName);
                                break;
                            case "Заголовок 1":
                                p.set_Style(header1);
                                break;
                            case "Заголовок 2":
                                p.set_Style(header2);
                                break;
                            case "Заголовок 3":
                                p.set_Style(header3);
                                break;
                            case "Заголовок 4":
                                p.set_Style(header4);
                                break;
                            case "Заголовок 5":
                                p.set_Style(header5);
                                break;
                        }
                    }

                    doc.Save();
                    doc.Close();
                    docTemplate.Close();
                    app.Quit(ref missingObj, ref  missingObj, ref missingObj);
                }
                else
                    throw new ReportException("Ненайден документ Word, содержащий стили le..", 1);

            }
            catch (ReportException re)
            {
                if (re.getPrintByte() == (byte)1)
                {
                    printPackageTree(re.getRepository().GetPackageByID(1));

                }

                re.eaClose();
                re.print(errWriter);
                app.Quit(ref missingObj, ref  missingObj, ref missingObj);
                Environment.Exit(1);
            }
            catch (Exception e)
            {
                errWriter.WriteLine(e.GetBaseException());
                app.Quit(ref missingObj, ref  missingObj, ref missingObj);
                Environment.Exit(2);
            }
        }

        //Возвращает ID пакета, вход - упорядоченный массив из наименований пакетов (путь к искомому пакету)
        public static int getPackageId(String[] list, EA.Repository r)
        {
            EA.Package package = r.GetPackageByID(1); //root package
            bool flag; //Найден ли очередной пакет
            
            for (short i = 1; i < list.Length; i++)
            {
                flag = false;
                for (short j = 0; j < package.Packages.Count; j++)
                {
                    if (package.Packages.GetAt(j).Name == list[i])
                    {
                        package = package.Packages.GetAt(j);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                    throw new ReportException("Не найден один из указанных пакетов, проверьте правильность введённого пути.", 1, r, 1);
            }

            return package.PackageID;
        }

        //Печатает список пакетов в модели.
        public static void printPackageTree(EA.Package package, String buffer = null)
        {
            EA.Package pp;          
            String temp;

            if (buffer == null)
                temp = package.Name;
            else
                temp = buffer;

            for (short i = 0; i < package.Packages.Count; i++)
            {
                pp = package.Packages.GetAt(i);
                Console.WriteLine(temp + "~" + pp.Name);

                if (pp.Packages.Count > 0)                  
                    printPackageTree(pp, temp + "~" + pp.Name);
            }

        }

        //Вывод Help-a в stdOut
        public static void printHelp()
        {
            Console.WriteLine("EAReport 1.0 Roman Letunovskii (c)");
            Console.WriteLine("Программа EAReport создана для быстрого получения документации в формате docx по существующему шаблону.");
            Console.WriteLine("");
            Console.WriteLine("Помимо аргумента вызывающего help принимается также массив из 5-и аргументов: ");
            Console.WriteLine("0 - Наименование файла модели (например - Каскад.eap).");
            Console.WriteLine("1 - Наименование выходного файла (например - report.docx).");
            Console.WriteLine("2 - Путь к документируемому пакету, разделитель ~ (например - ПК \"Каскад\"~Клиент~АРМ ОПК).");
            Console.WriteLine("3 - Наименование шаблона документации (например - sequential).");
            Console.WriteLine("4 - Путь к документу Word, в котором содержатся стили le..");
            Console.WriteLine("");
            Console.WriteLine("Программа возвращает коды ошибок: ");
            Console.WriteLine("0 - Программа отработала без ошибок.");
            Console.WriteLine("1 - Программа отработала с не критичными ошибками.");
            Console.WriteLine("2 - Программа не смогла выполниться.");
        }
    }

    //Класс для обработки ошибок документации
    public class ReportException : Exception 
    {
        String message;
        int code;
        EA.Repository r;
        byte p;

        public byte getPrintByte()
        {
            return p;
        }

        public EA.Repository getRepository()
        {
            return r;
        }

        public ReportException(String message, int code)
        {
            this.message = message;
            this.code = code;
        }

        public ReportException(String message, int code, EA.Repository r)
        {
            this.message = message;
            this.code = code;
            this.r = r;
        }

        public ReportException(String message, int code, EA.Repository r, byte b)
        {
            this.message = message;
            this.code = code;
            this.r = r;
            this.p = b;
        }

        public void print(TextWriter tw)
        {
            tw.WriteLine(code + " : " + message);
        }

        public void eaClose()
        {
            if (r != null)
            {
                r.CloseFile();
                r.Exit();
            }
        }
    }
}
