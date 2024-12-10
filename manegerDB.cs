using Aspose.Cells;
using Aspose.Cells.Charts;
using System.ComponentModel.Design;
using System.Net;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using static System.Runtime.InteropServices.JavaScript.JSType;

public class manegerBD
{
    private Logger logger;
    private string PathToXlsFile;
    private BD hotel;

    public manegerBD(string path, Logger logger)
    {
        PathToXlsFile = path;
        this.logger = logger;
        hotel = new BD(PathToXlsFile);
    }

    public void OutputBD() // вывод БД
    {
        string chare = new string('_', 120);
        Console.WriteLine("Таблица - Клиенты:");

        Console.WriteLine(chare);
        Console.WriteLine($"Код | Фамилия              | Имя        | Отчество        | Место жительства");
        foreach (var client in hotel.clients)
        {
            Console.WriteLine(client.ToString());
        }
        Console.WriteLine(chare);
        Console.WriteLine("Таблица - Номера:");

        Console.WriteLine(chare);
        Console.WriteLine($"Код   |Этаж |Места| Стоимость| Категория");
        foreach (var number in hotel.numbers)
        {
            Console.WriteLine(number.ToString());
        }
        Console.WriteLine(chare);
        Console.WriteLine($"Таблицы - Бронь:");

        Console.WriteLine(chare);
        Console.WriteLine($"Код   |Код клиента | Код номера | Дата брони      | Дата заезда     | Дата выезда");
        foreach (var reserv in hotel.reservation)
        {
            Console.WriteLine(reserv.ToString());
        }
        Console.WriteLine();
        logger.Log("БД выведена");
    }  

    public void CorrectionElement() // Корректировка элементов БД 
    {
        Console.WriteLine("Введите номер таблицы корректировки: ");
        Console.WriteLine("1. Клиенты");
        Console.WriteLine("2. Номера");
        Console.WriteLine("3. Бронирования");
        string ch = Console.ReadLine();

        Console.WriteLine("Введите ключ элемента который хотите изменить");
        if (uint.TryParse(Console.ReadLine(), out uint key))
        {
            switch (ch)
            {
                case "1":
                    CorrectionElementsClients(key);
                    break;
                case "2":
                    CorrectionElementsNumbers(key);
                    break;
                case "3":
                    CorrectionElementsReservation(key);
                    break;
                default:
                    Console.WriteLine("asdasd");
                    break;
            }
        }
        else 
        { 
            Console.WriteLine("NaN"); 
        }
       
    }

    public void AddElement() // добавление элементов БД
    {
        Console.WriteLine("ВВедите номер таблицы для добавления");
        Console.WriteLine("1. Клиенты");
        Console.WriteLine("2. Номера");
        Console.WriteLine("3. Бронирования");
        string ch = Console.ReadLine();

        switch (ch)
        {
            case "1":
                AddClients();
                break;
            case "2":
                AddNumber();
                break;
            case "3":
                AddReservation();
                break;
            default:
                Console.WriteLine();
                break;
        }
    }

    public void DeleteElement() // удаление элементов БД
    {
        Console.WriteLine("Введите номер таблицы из которой хотите удалить");
        Console.WriteLine("1. Клиенты");
        Console.WriteLine("2. Номера");
        Console.WriteLine("3. Бронирования");
        string ch = Console.ReadLine();

        Console.WriteLine("Введите ключ таблицы для удаления элемента");
        if(uint.TryParse(Console.ReadLine(), out uint key))
        switch (ch)
        {
            case "1":
                DeleteElementFromClients(key);
                break;
            case "2":
                DeleteElementFromNumbers(key);
                break;
            case "3":
                DeleteElementFromReservation(key);
                break;
            default:
                Console.WriteLine();
                break;
        }
        else
        {
            Console.WriteLine("NaN");
        }
    }

    // вспомогательные методы для публичных методов выше
    private void DeleteElementFromReservation(uint key)
    {
        var deleteReservation = hotel.reservation.FirstOrDefault(b => b.id == key);
        logger.Log($"Удалена запись из таблицы Бронирование: {deleteReservation.ToString()}");
        if (deleteReservation != null)
        {
            var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == deleteReservation.numbersId);
            var deleteClients = hotel.clients.FirstOrDefault(c => c._id == deleteReservation.clientsId);

            hotel.reservation.Remove(deleteReservation);
            hotel.numbers.Remove(deleteNumbers);
            hotel.clients.Remove(deleteClients);
        }
        else
        {
            Console.WriteLine("такого элементта не существует");
        }
    }

    private void DeleteElementFromClients(uint clientsKey)
    {
        var deleteClients = hotel.clients.FirstOrDefault(c => c._id == clientsKey);
        logger.Log($"Удалена запись из таблицы Клиенты: {deleteClients.ToString()}");
        if (deleteClients != null)
        {
            var deleteReservation = hotel.reservation.FirstOrDefault(r => r.clientsId == deleteClients._id);
            var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == deleteReservation.numbersId);

            hotel.reservation.Remove(deleteReservation);
            hotel.numbers.Remove(deleteNumbers);
            hotel.clients.Remove(deleteClients);
        }
    }

    private void DeleteElementFromNumbers(uint numberKey)
    {
        var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == numberKey);
        logger.Log($"Удалена запись из таблицы Номера: {deleteNumbers.ToString()}");
        if (deleteNumbers != null)
        {
            var deleteReservation = hotel.reservation.FirstOrDefault(r => r.numbersId == numberKey);
            var deleteClients = hotel.clients.FirstOrDefault(c => c._id == deleteReservation.clientsId);

            hotel.reservation.Remove(deleteReservation);
            hotel.numbers.Remove(deleteNumbers);
            hotel.clients.Remove(deleteClients);
        }
        
    }

    private void CorrectionElementsClients(uint key)
    {
        var client = hotel.clients.FirstOrDefault(c => c._id == key);
        if (client == null)
        {
            return;
        }

        Console.WriteLine("Введите новую фамилию(чтобы не менять оставьте пустым)");
        string newSecName = Console.ReadLine();
        if (!string.IsNullOrEmpty(newSecName)) client._firstName = newSecName;

        Console.WriteLine("Введите новое имя");
        string newfirstName = Console.ReadLine();
        if (!string.IsNullOrEmpty(newfirstName)) client._secondName = newfirstName;

        Console.WriteLine("введите новое отчество");
        string newPatr = Console.ReadLine();
        if (!string.IsNullOrEmpty(newPatr)) client._patronymic = newPatr;

        Console.WriteLine("Введите новый адрес");
        string newAdress = Console.ReadLine();
        if (!string.IsNullOrEmpty(newAdress)) client._adres = newAdress;
        logger.Log($"изменена запись в таблице Клиенты измененная запись: {client.ToString()}");
    }

    private void CorrectionElementsNumbers(uint key)
    {
        var number = hotel.numbers.FirstOrDefault(n => n._id == key);
        if (number == null)
        {
            return;
        }

        Console.WriteLine("Введите новый этаж: ");
        string newFloor = Console.ReadLine();
        if (ushort.TryParse(newFloor, out ushort newFloorInput)) number._floor = newFloorInput;

        Console.WriteLine("Введите новое количество мест: ");
        string newPlace = Console.ReadLine();
        if (ushort.TryParse(newPlace, out ushort newPlacesInput)) number._numbeerOfplaces = newPlacesInput;

        Console.WriteLine("Введите новую цену: ");
        string newPrice = Console.ReadLine();
        if (uint.TryParse(newPrice, out uint newPriceInput)) number._price = newPriceInput;

        Console.WriteLine("ВВедите новую категорию номера: ");
        string newCtegory = Console.ReadLine();
        if (ushort.TryParse(newCtegory, out ushort newCtegoryInput)) number._category = newCtegoryInput;
        logger.Log($"изменена запись в таблице Номера измененная запись: {number.ToString()}");
    }

    private void CorrectionElementsReservation(uint key)
    {
        var reservation = hotel.reservation.FirstOrDefault(r => r.id == key);
        if (reservation == null) { return; }

        Console.WriteLine("Введите новую дату бронирования вида дд.мм.гггг");
        string newReservationDate = Console.ReadLine();
        if (!string.IsNullOrEmpty(newReservationDate))
        {
            reservation.ReservationData = new CustomDate(newReservationDate);
        }

        Console.WriteLine("Введите новую дату заезда");
        string newReservationChekInDate = Console.ReadLine();
        if (!string.IsNullOrEmpty(newReservationChekInDate))
        {
            reservation.ChekInData = new CustomDate(newReservationChekInDate);
        }

        Console.WriteLine("Введите новую дату выезда");
        string newDepartureData = Console.ReadLine();
        if (!string.IsNullOrEmpty(newDepartureData))
        {
            reservation.departureDate = new CustomDate(newDepartureData);
        }
        logger.Log($"изменена запись в таблице Резервация измененная запись: {reservation.ToString()}");
    }

    private void AddClients()
    {
        Console.WriteLine("Введите новую фамилию");
        string newSecName = Console.ReadLine();


        Console.WriteLine("Введите новое имя");
        string newfirstName = Console.ReadLine();

        Console.WriteLine("введите новое отчество");
        string newPatr = Console.ReadLine();

        Console.WriteLine("Введите новый адрес");
        string newAdress = Console.ReadLine();

        if (!string.IsNullOrEmpty(newAdress) || !string.IsNullOrEmpty(newPatr) || !string.IsNullOrEmpty(newfirstName) || string.IsNullOrEmpty(newSecName))
        {

            hotel.clients.Add(new Сlients(hotel.clients.Max(c => c._id) + 1, newSecName, newfirstName, newPatr, newAdress));
        }
        logger.Log($"Добавлена новая запись в таблицу клиенты с полями {newSecName}, {newfirstName}, {newPatr}, {newAdress}");
    }

    private void AddNumber()
    {
        Console.WriteLine("Введите новый этаж: ");
        string newFloor = Console.ReadLine();

        Console.WriteLine("Введите новое количество мест: ");
        string newPlace = Console.ReadLine();

        Console.WriteLine("Введите новую цену: ");
        string newPrice = Console.ReadLine();

        Console.WriteLine("ВВедите новую категорию номера: ");
        string newCtegory = Console.ReadLine();

        if (ushort.TryParse(newFloor, out ushort NewFloorOutput) && ushort.TryParse(newPlace, out ushort NewPlaceOutput) && uint.TryParse(newPrice, out uint NewPriceOutput) && ushort.TryParse(newCtegory, out ushort newCategoryOutput))
        {
            hotel.numbers.Add(new Numbers(hotel.numbers.Max(c => c._id) + 1, NewFloorOutput, NewPlaceOutput, NewPlaceOutput, newCategoryOutput));
        }
        logger.Log($"Добавлена новая запись в таблицу номера с полями {newFloor}, {newPlace}, {newPrice}, {newCtegory}");
    }

    private void AddReservation()
    {
        Console.WriteLine("Введите новую дату бронирования вида дд.мм.гггг");
        string newReservationDate = Console.ReadLine();

        Console.WriteLine("Введите новую дату заезда");
        string newReservationChekInDate = Console.ReadLine();

        Console.WriteLine("Введите новую дату выезда");
        string newDepartureData = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(newReservationDate) && !string.IsNullOrEmpty(newReservationDate) && !string.IsNullOrEmpty(newDepartureData))
        {
            hotel.reservation.Add(new Reservation(
                hotel.reservation.Max(r => r.id) + 1,
                hotel.clients.Max(c => c._id) + 1,
                hotel.numbers.Max(n => n._id) + 1,
                new CustomDate(newReservationDate),
                new CustomDate(newReservationChekInDate),
                new CustomDate(newDepartureData))
                );
        }
        logger.Log($"Добавлена новая запись в таблицу Бронирование и зависимости с полями {newReservationDate}, {newReservationChekInDate}, {newDepartureData}");
        AddClients();
        AddNumber();
    }

    //запросы
    public void GetTotalClients()
    {
        Console.WriteLine(hotel.clients.Count());
    }
    public void GetReservedNumbers()
    {
        
        var reservedNumbers = hotel.reservation
            .Select(r => hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._id)
            .Where(n => n != null)
            .ToList();

        foreach (var reservedNumber in reservedNumbers)
        {
            Console.Write(reservedNumber + " ");
        }
    }

    public void GetClientsWithReservation()
    {
        var clientsWithReservation = hotel.reservation
            .Select (r => new
            {
                ClinetName = hotel.clients.FirstOrDefault(c => c._id == r.clientsId)?._secondName,
                RoomsNumber = hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._id,
                Price = hotel.numbers.FirstOrDefault(n => n._id==r.numbersId)?._price
            }).Where(res => res.ClinetName != null && res.RoomsNumber != null)
            .Select(res => $"{res.ClinetName, -15}| {res.RoomsNumber, -10}| {res.Price, - 5}|").ToList();

        Console.WriteLine();
        Console.WriteLine("клиент         |id комнаты |цена  |");
        foreach (var client in clientsWithReservation)
        {
            Console.WriteLine(client);
        }
    }

    public void GetRevenueForCity(string city) // запрос на стоимость брони для клиентов проживающих в 1 городе
    {
        decimal totalCost = hotel.reservation
            .Where(r => hotel.clients.FirstOrDefault(c => c._id == r.clientsId)?._adres == city).Sum(r => hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._price ?? 0);

        Console.WriteLine(totalCost);
    }


    public void SaveChagesInXls()
    {
        Workbook wb = new Workbook(PathToXlsFile);

        Worksheet wsClients = wb.Worksheets["Клиенты"];
        Worksheet wsReservation = wb.Worksheets["Бронирование"];
        Worksheet wsNumbers = wb.Worksheets["Номера"];

        for(int i = 0; i < hotel.clients.Count; i++)
        {
            var client = hotel.clients[i];
            wsClients.Cells[$"A{i + 2}"].PutValue(client._id);
            wsClients.Cells[$"B{i + 2}"].PutValue(client._firstName);
            wsClients.Cells[$"C{i + 2}"].PutValue(client._secondName);
            wsClients.Cells[$"D{i + 2}"].PutValue(client._patronymic);
            wsClients.Cells[$"E{i + 2}"].PutValue(client._adres);
        }

        for(int i = 0; i < hotel.numbers.Count; i++)
        {
            var number = hotel.numbers[i];
            wsNumbers.Cells[$"A{i + 2}"].PutValue(number._id);
            wsNumbers.Cells[$"B{i + 2}"].PutValue(number._floor);
            wsNumbers.Cells[$"C{i + 2}"].PutValue(number._numbeerOfplaces);
            wsNumbers.Cells[$"D{i + 2}"].PutValue(number._price);
            wsNumbers.Cells[$"E{i + 2}"].PutValue(number._category);
        }

        for(int i = 0; i < hotel.reservation.Count; i++)
        {
            var reservation = hotel.reservation[i];
            wsReservation.Cells[$"A{i + 2}"].PutValue(reservation.id);
            wsReservation.Cells[$"B{i + 2}"].PutValue(reservation.clientsId);
            wsReservation.Cells[$"C{i + 2}"].PutValue(reservation.numbersId);
            wsReservation.Cells[$"D{i + 2}"].PutValue(reservation.ReservationData);
            wsReservation.Cells[$"E{i + 2}"].PutValue(reservation.ChekInData);
            wsReservation.Cells[$"F{i + 2}"].PutValue(reservation.departureDate);
        }
        wb.Save(PathToXlsFile);

    }
}