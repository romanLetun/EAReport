using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace EAReport
{
    class EAReport
    {
        const char delimiter = '~'; //Разделитель пакетов в Enterprise Architect (далее EA).
        const String fileDelimiter = "\\"; //Разделитель.
        const String format = ".docx"; //Формат выходного документа.
        const String emptyString = "^p^p"; //Пустая строка в Word.     
        const String help = "h"; //Параметр для вызова Help
        const String anotherHelp = "?"; //Параметр для вызова Help   
        const String sixthArg = "Y"; // Определяет удалять ли пустые строки в выходном документе. По умолчанию не удалять.
        const String stylesDefault = "leCode:lePictureName,leUsual:leTableName,Заголовок 1:leHeader1,Заголовок 2:leHeader2,Заголовок 3:leHeader3,Заголовок 4:leHeader4,Заголовок 5:leHeader5"; //замена стилей по умолчанию
        //Сообщения об ошибках.
        const String withoutArgs = "Запуск без параметров. 83B1A393-D7EC-4405-8BBE-A845371EBF31";
        const String incorrectArgsCount = "Не верное количество параметров. D25A7717-4F55-4218-A3B5-DEA69B7B8904";
        const String fileNotFound = "Файл модели не найден. A1B4BA36-0636-47BF-A25F-7B76ABCADAA2";
        const String incorrect = "Наименование выходного файла должно соответствовать шаблону \\w+\\.docx. C70E8CCC-4CDF-4FE2-8CF0-CC41938947BE";
        const String incorrectPath = "Путь к документируемому пакету должен содержать хотя бы один символ \"~\". 8141B24E-1B2E-484F-A016-1A445D199384";
        const String templateNotFound = "Ненайден документ Word, содержащий стили le.. FF73C14B-04C4-4E35-845B-188D03FA48D2";
        const String packageNotFound = "Не найден один из указанных пакетов, проверьте правильность введённого пути. AA64B99C-BF24-4BDA-BFA9-47336F7CF6D0";

        public static SortedDictionary<String, String> styleNames = new SortedDictionary<string, string>();

        static void Main(string[] args)
        {    
            int packageID; //ID пакета в EA.
            String[] list; //
            String template; //Наименование шаблона документации в EA.
            String path = null; //Выходной документ.
            String docTemplatePath; // Путь к документу word, в котором находятся стили для замены.
            TextWriter errWriter = Console.Error; //Вывод в stderror.  
            StreamReader reader; // Чтение входного файла переданного в 7-ом параметре.
            Object missingObj = System.Reflection.Missing.Value;
            Word.Application app = new Word.Application();
            Object missing = Type.Missing;

            try
            {               
                if (args.Length < 1)
                {
                    throw new ReportException(withoutArgs, 1);
                }

                if (args[0] == help || args[0] == anotherHelp)
                {
                    printHelp();
                    Environment.Exit(1);
                }

                else
                    if (args.Length >= 5)
                    {
                        EA.Repository r = new EA.Repository();
                        EA.Project project = new EA.Project();

                        if (args.Length > 6)
                        {
                            reader = new StreamReader(args[6], Encoding.Default);//File.OpenText(args[6]);
                            styles(reader);
                            reader.Close();
                        }
                        else
                            styles();

                        //Открываем файл модели
                        if (!args[0].Contains(fileDelimiter))
                            args[0] = Environment.CurrentDirectory + fileDelimiter + args[0];

                        if (File.Exists(args[0]))
                            r.OpenFile(args[0]);
                        else
                            throw new ReportException(fileNotFound, 1);

                        //Выходной файл
                        if (args[1].Substring(args[1].Length - 5) == format)
                            path = Environment.CurrentDirectory + fileDelimiter + args[1];
                        else
                            throw new ReportException(incorrect, 1);

                        //Путь к документируемому пакету
                        list = args[2].Split(delimiter);

                        if (list.Length < 2)
                            throw new ReportException(incorrectPath, 1);

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
                        throw new ReportException(incorrectArgsCount, 1);
                
                if (args[4].Contains(fileDelimiter))
                    docTemplatePath = args[4];
                else
                    docTemplatePath = Environment.CurrentDirectory + fileDelimiter + args[4];

                if (File.Exists(path) && File.Exists(docTemplatePath))
                {                    
                    Word.Document doc = app.Documents.Open(path);
                    
                    if (args.Length > 5)
                    {
                        if (args[5].Equals(sixthArg))
                        {
                            Word.Find find = app.Selection.Find;
                            find.Text = emptyString;
                            find.Replacement.Text = emptyString.Substring(0, 2);
                            Object wrap = Word.WdFindWrap.wdFindContinue;
                            Object replace = Word.WdReplace.wdReplaceAll;
                            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false, MatchWildcards: false, MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true, Wrap: wrap, Format: false, ReplaceWith: missing, Replace: replace);
                        }
                    }

                    Word.Document docTemplate = app.Documents.Open(docTemplatePath);

                    foreach (Word.Paragraph p in doc.Paragraphs)
                    {
                        if (styleNames.ContainsKey(p.get_Style().NameLocal))
                        {
                            foreach (Word.Style s in docTemplate.Styles)
                            {
                                if (s.NameLocal == styleNames[p.get_Style().NameLocal])
                                {
                                    p.set_Style(s);
                                    break;
                                }
                            }
                        }
                    }

                    doc.Save();
                    doc.Close();
                    docTemplate.Close();
                    app.Quit(ref missingObj, ref  missingObj, ref missingObj);
                }
                else
                    throw new ReportException(templateNotFound, 1);

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
                    throw new ReportException(packageNotFound, 1, r, 1);
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
                Console.WriteLine(temp + delimiter + pp.Name);

                if (pp.Packages.Count > 0)
                    printPackageTree(pp, temp + delimiter + pp.Name);
            }

        }

        //Вывод Help-a в stdOut
        public static void printHelp()
        {
            Console.WriteLine("EAReport 1.1 (c) Roman Letunovskii, 2015");
            Console.WriteLine("Программа EAReport создана для быстрого получения документации в формате docx по существующему шаблону.");
            Console.WriteLine("");
            Console.WriteLine("Помимо параметра вызывающего help принимается также массив из 7-и параметров: ");
            Console.WriteLine("1 - Наименование файла модели (например - Каскад.eap).");
            Console.WriteLine("2 - Наименование выходного файла (например - report.docx).");
            Console.WriteLine("3 - Путь к документируемому пакету, разделитель ~ (например - ПК \"Каскад\"~Клиент~АРМ ОПК).");
            Console.WriteLine("4 - Наименование шаблона документации (например - sequential).");
            Console.WriteLine("5 - Путь к документу Word, в котором содержатся стили le..");
            Console.WriteLine("6 - Нужно ли удалять пустые строки в выходном документе \"Y\" - удалять. Не обязательный параметр.");
            Console.WriteLine("7 - Необязательный параметр файл с перечнем наименований стилей, которые необходимо заменить в выходном документе");
            Console.WriteLine("и наименований стилей, на которые их надо заменить одной строкой в формате:");
            Console.WriteLine("Наименование заменяемого стиля:Наименование стиля для замены");
            Console.WriteLine("Разделитель между наименованиями \":\", разделитель между парами стилей \",\"");
            Console.WriteLine("По умолчанию заменяются стили leCode:lePictureName,leUsual:leTableName,Заголовок 1:leHeader1,Заголовок 2:leHeader2,Заголовок 3:leHeader3,Заголовок 4:leHeader4,Заголовок 5:leHeader5");
            Console.WriteLine("Кодировка файла - 1251");
            Console.WriteLine("");
            Console.WriteLine("Параметры с имененм файла можно вводить с полным путём к файлу или же только имя файла, в таком случае он должен находиться в текущем каталоге.");
            Console.WriteLine("");
            Console.WriteLine("Сообщения об ошибках выводятся в strError.");
            Console.WriteLine("");
            Console.WriteLine("Программа возвращает коды ошибок: ");
            Console.WriteLine("0 - Программа отработала без ошибок.");
            Console.WriteLine("1 - Программа отработала с не критичными ошибками.");
            Console.WriteLine("2 - Программа не смогла выполниться.");
        }

        //Сохраняет наименования стилей из stdout в styleNames.
        public static void styles(StreamReader reader)
        {
            String sNames = reader.ReadLine();
            String[] temp;

            if (styleNames == null)
                temp = stylesDefault.Split(',');//Console.ReadLine().Split(',');
            else
                temp = sNames.Split(',');

            foreach (String s in temp)
            {                
                styleNames.Add(s.Split(':')[0], s.Split(':')[1]);
            }
        }

        public static void styles()
        {
            String[] temp;
        
            temp = stylesDefault.Split(',');//Console.ReadLine().Split(',');

            foreach (String s in temp)
            {
                styleNames.Add(s.Split(':')[0], s.Split(':')[1]);
            }
        }
    }

    //Класс для обработки ошибок документации
    public class ReportException : Exception 
    {
        String message; //Сообщение об ошибке.
        int code; //Код ошибки
        EA.Repository r; 
        byte p;
        const String errorMessageDelimiter = " : ";

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
            tw.WriteLine(code + errorMessageDelimiter + message);
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
