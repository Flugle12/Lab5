using System.Diagnostics;
using System.IO;

public class Program
{
    public static void Main()
    {
        string path = "";

        try
        {
            Console.WriteLine("Вы хотите протоколировать ваши действия в новый файл или в существующий?");
            Console.WriteLine("1. Новый файл");
            Console.WriteLine("2. Существующий файл");
            string ch = Console.ReadLine();

            switch (ch)
            {
                case "1":
                    Console.WriteLine("Введите название нового файла");
                    path = Console.ReadLine();
                    break;
                case "2":
                    Console.WriteLine("Укажите название этого файла");
                    path = Console.ReadLine();
                    break;
                default:
                    Console.WriteLine("Неверный выбор.");
                    return;
            }

            Logger logger = new Logger(path);
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

                var key = Console.ReadKey(true);

                if (key.Key == ConsoleKey.Escape)
                {
                    Console.WriteLine("Выход из программы.");
                    break;
                }

                try
                {
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
                catch (Exception ex)
                {
                    Console.WriteLine($"Произошла ошибка: {ex.Message}");
                }
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
        }
    }
}
