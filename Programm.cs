
using System.Diagnostics;
using System.IO;
using System.Net.Http.Headers;

public class Program
{
    public static void Main()
    {
        Console.WriteLine("вы хотите протоколировать ваши действия в новый файл или в преведущий?");
        Console.WriteLine("1. Новый файл");
        Console.WriteLine("2. Существующий файл");
        string ch = Console.ReadLine();

        string Path = "";
        switch (ch)
        {
            case "1":
                Console.WriteLine("Введите название нового файла");
                Path = Console.ReadLine();
                break;
            case "2":
                Console.WriteLine("укажите название этого файла");
                Path = Console.ReadLine(); ;
                break;
            default: Console.WriteLine("a");
                break;
        }
        Logger logger = new Logger(Path);


        manegerBD meneger = new manegerBD("C:\\Users\\User\\Desktop\\LR5 Excel Files\\LR5-var9.xls", logger);

        Console.WriteLine("Нажмите 'Esc' для выхода из программы.");
        while (true)
        {
            Console.WriteLine("\nВыберите действие:");
            Console.WriteLine("1. Вывести базу данных");
            Console.WriteLine("2. Добавить элемент");
            Console.WriteLine("3. Удалить элемент");
            Console.WriteLine("4. Корректировать элемент");
            Console.WriteLine("5. Сохранить изменения");
            Console.WriteLine("Нажмите 'Esc' для выхода.");

            var key = Console.ReadKey(true); // Считываем нажатую клавишу

            if (key.Key == ConsoleKey.Escape)
            {
                Console.WriteLine("Выход из программы.");
                break; // Выход из цикла
            }

            switch (key.KeyChar)
            {
                case '1':
                    meneger.OutputBD();
                    break;
                case '2':
                    meneger.AddElement();
                    break;
                case '3':
                    meneger.DeleteElement();
                    break;
                case '4':
                    meneger.CorrectionElement();
                    break;
                case '5':
                    meneger.SaveChagesInXls();
                    Console.WriteLine("Изменения сохранены.");
                    break;
                default:
                    Console.WriteLine("Неверный выбор. Пожалуйста, попробуйте снова.");
                    break;
            }
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = Path,
            UseShellExecute = true
        });

    }
}

