using Aspose.Cells;
using System;
using System.Linq;

public class manegerBD
{
    private Logger logger;
    private string PathToXlsFile;
    private BD hotel;

    public manegerBD(string path, Logger logger)
    {
        PathToXlsFile = path;
        this.logger = logger;
        hotel = new BD(PathToXlsFile, this.logger);
    }

    public void OutputBD() // вывод БД
    {
        try
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
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при выводе БД: {ex.Message}");
            logger.Log($"Ошибка при выводе БД: {ex.Message}");
        }
    }

    public void CorrectionElement() // Корректировка элементов БД 
    {
        try
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
                        Console.WriteLine("Неверный выбор.");
                        break;
                }
            }
            else
            {
                Console.WriteLine("Некорректный ввод ключа.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при корректировке элемента: {ex.Message}");
            logger.Log($"Ошибка при корректировке элемента: {ex.Message}");
        }
    }

    public void AddElement() // добавление элементов БД
    {
        try
        {
            Console.WriteLine("Введите номер таблицы для добавления");
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
                    Console.WriteLine("Неверный выбор.");
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при добавлении элемента: {ex.Message}");
            logger.Log($"Ошибка при добавлении элемента: {ex.Message}");
        }
    }

    public void DeleteElement() // удаление элементов БД
    {
        try
        {
            Console.WriteLine("Введите номер таблицы из которой хотите удалить");
            Console.WriteLine("1. Клиенты");
            Console.WriteLine("2. Номера");
            Console.WriteLine("3. Бронирования");
            string ch = Console.ReadLine();

            Console.WriteLine("Введите ключ таблицы для удаления элемента");
            if (uint.TryParse(Console.ReadLine(), out uint key))
            {
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
                        Console.WriteLine("Неверный выбор.");
                        break;
                }
            }
            else
            {
                Console.WriteLine("Некорректный ввод ключа.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при удалении элемента: {ex.Message}");
            logger.Log($"Ошибка при удалении элемента: {ex.Message}");
        }
    }

    // Вспомогательные методы для публичных методов выше
    private void DeleteElementFromReservation(uint key)
    {
        try
        {
            var deleteReservation = hotel.reservation.FirstOrDefault(b => b.id == key);
            if (deleteReservation != null)
            {
                logger.Log($"Удалена запись из таблицы Бронирование: {deleteReservation.ToString()}");
                var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == deleteReservation.numbersId);
                var deleteClients = hotel.clients.FirstOrDefault(c => c._id == deleteReservation.clientsId);

                hotel.reservation.Remove(deleteReservation);
                if (deleteNumbers != null) hotel.numbers.Remove(deleteNumbers);
                if (deleteClients != null) hotel.clients.Remove(deleteClients);
            }
            else
            {
                Console.WriteLine("Такого элемента не существует");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при удалении из бронирования: {ex.Message}");
            logger.Log($"Ошибка при удалении из бронирования: {ex.Message}");
        }
    }

    private void DeleteElementFromClients(uint clientsKey)
    {
        try
        {
            var deleteClients = hotel.clients.FirstOrDefault(c => c._id == clientsKey);
            if (deleteClients != null)
            {
                logger.Log($"Удалена запись из таблицы Клиенты: {deleteClients.ToString()}");
                var deleteReservation = hotel.reservation.FirstOrDefault(r => r.clientsId == deleteClients._id);
                var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == deleteReservation?.numbersId);

                hotel.reservation.Remove(deleteReservation);
                if (deleteNumbers != null) hotel.numbers.Remove(deleteNumbers);
                hotel.clients.Remove(deleteClients);
            }
            else
            {
                Console.WriteLine("Такого клиента не существует.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при удалении клиента: {ex.Message}");
            logger.Log($"Ошибка при удалении клиента: {ex.Message}");
        }
    }

    private void DeleteElementFromNumbers(uint numberKey)
    {
        try
        {
            var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == numberKey);
            if (deleteNumbers != null)
            {
                logger.Log($"Удалена запись из таблицы Номера: {deleteNumbers.ToString()}");
                var deleteReservation = hotel.reservation.FirstOrDefault(r => r.numbersId == numberKey);
                var deleteClients = hotel.clients.FirstOrDefault(c => c._id == deleteReservation?.clientsId);

                hotel.reservation.Remove(deleteReservation);
                if (deleteClients != null) hotel.clients.Remove(deleteClients);
                hotel.numbers.Remove(deleteNumbers);
            }
            else
            {
                Console.WriteLine("Такого номера не существует.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при удалении номера: {ex.Message}");
            logger.Log($"Ошибка при удалении номера: {ex.Message}");
        }
    }

    private void CorrectionElementsClients(uint key)
    {
        try
        {
            var client = hotel.clients.FirstOrDefault(c => c._id == key);
            if (client == null)
            {
                Console.WriteLine("Клиент не найден.");
                return;
            }

            Console.WriteLine("Введите новую фамилию (чтобы не менять, оставьте пустым)");
            string newSecName = Console.ReadLine();
            if (!string.IsNullOrEmpty(newSecName)) client._firstName = newSecName;

            Console.WriteLine("Введите новое имя (чтобы не менять, оставьте пустым)");
            string newfirstName = Console.ReadLine();
            if (!string.IsNullOrEmpty(newfirstName)) client._secondName = newfirstName;

            Console.WriteLine("Введите новое отчество (чтобы не менять, оставьте пустым)");
            string newPatr = Console.ReadLine();
            if (!string.IsNullOrEmpty(newPatr)) client._patronymic = newPatr;

            Console.WriteLine("Введите новый адрес (чтобы не менять, оставьте пустым)");
            string newAdress = Console.ReadLine();
            if (!string.IsNullOrEmpty(newAdress)) client._adres = newAdress;

            logger.Log($"Изменена запись в таблице Клиенты: {client.ToString()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при корректировке клиента: {ex.Message}");
            logger.Log($"Ошибка при корректировке клиента: {ex.Message}");
        }
    }

    private void CorrectionElementsNumbers(uint key)
    {
        try
        {
            var number = hotel.numbers.FirstOrDefault(n => n._id == key);
            if (number == null)
            {
                Console.WriteLine("Номер не найден.");
                return;
            }

            Console.WriteLine("Введите новый этаж (чтобы не менять, оставьте пустым): ");
            string newFloor = Console.ReadLine();
            if (ushort.TryParse(newFloor, out ushort newFloorInput)) number._floor = newFloorInput;

            Console.WriteLine("Введите новое количество мест (чтобы не менять, оставьте пустым): ");
            string newPlace = Console.ReadLine();
            if (ushort.TryParse(newPlace, out ushort newPlacesInput)) number._numbeerOfplaces = newPlacesInput;

            Console.WriteLine("Введите новую цену (чтобы не менять, оставьте пустым): ");
            string newPrice = Console.ReadLine();
            if (uint.TryParse(newPrice, out uint newPriceInput)) number._price = newPriceInput;

            Console.WriteLine("Введите новую категорию номера (чтобы не менять, оставьте пустым): ");
            string newCtegory = Console.ReadLine();
            if (ushort.TryParse(newCtegory, out ushort newCtegoryInput)) number._category = newCtegoryInput;

            logger.Log($"Изменена запись в таблице Номера: {number.ToString()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при корректировке номера: {ex.Message}");
            logger.Log($"Ошибка при корректировке номера: {ex.Message}");
        }
    }

    private void CorrectionElementsReservation(uint key)
    {
        try
        {
            var reservation = hotel.reservation.FirstOrDefault(r => r.id == key);
            if (reservation == null)
            {
                Console.WriteLine("Бронирование не найдено.");
                return;
            }

            Console.WriteLine("Введите новую дату бронирования (вида дд.мм.гггг, чтобы не менять, оставьте пустым):");
            string newReservationDate = Console.ReadLine();
            if (!string.IsNullOrEmpty(newReservationDate))
            {
                reservation.ReservationData = new CustomDate(newReservationDate);
            }

            Console.WriteLine("Введите новую дату заезда (чтобы не менять, оставьте пустым):");
            string newReservationChekInDate = Console.ReadLine();
            if (!string.IsNullOrEmpty(newReservationChekInDate))
            {
                reservation.ChekInData = new CustomDate(newReservationChekInDate);
            }

            Console.WriteLine("Введите новую дату выезда (чтобы не менять, оставьте пустым):");
            string newDepartureData = Console.ReadLine();
            if (!string.IsNullOrEmpty(newDepartureData))
            {
                reservation.departureDate = new CustomDate(newDepartureData);
            }

            logger.Log($"Изменена запись в таблице Резервация: {reservation.ToString()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при корректировке бронирования: {ex.Message}");
            logger.Log($"Ошибка при корректировке бронирования: {ex.Message}");
        }
    }

    private void AddClients()
    {
        try
        {
            Console.WriteLine("Введите новую фамилию:");
            string newSecName = Console.ReadLine();

            Console.WriteLine("Введите новое имя:");
            string newfirstName = Console.ReadLine();

            Console.WriteLine("Введите новое отчество:");
            string newPatr = Console.ReadLine();

            Console.WriteLine("Введите новый адрес:");
            string newAdress = Console.ReadLine();

            if (!string.IsNullOrEmpty(newSecName) && !string.IsNullOrEmpty(newfirstName))
            {
                hotel.clients.Add(new Сlients(hotel.clients.Max(c => c._id) + 1, newSecName, newfirstName, newPatr, newAdress));
                logger.Log($"Добавлена новая запись в таблицу клиенты с полями {newSecName}, {newfirstName}, {newPatr}, {newAdress}");
            }
            else
            {
                Console.WriteLine("Фамилия и имя не могут быть пустыми.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при добавлении клиента: {ex.Message}");
            logger.Log($"Ошибка при добавлении клиента: {ex.Message}");
        }
    }

    private void AddNumber()
    {
        try
        {
            Console.WriteLine("Введите новый этаж: ");
            string newFloor = Console.ReadLine();

            Console.WriteLine("Введите новое количество мест: ");
            string newPlace = Console.ReadLine();

            Console.WriteLine("Введите новую цену: ");
            string newPrice = Console.ReadLine();

            Console.WriteLine("Введите новую категорию номера: ");
            string newCtegory = Console.ReadLine();

            if (ushort.TryParse(newFloor, out ushort NewFloorOutput) &&
                ushort.TryParse(newPlace, out ushort NewPlaceOutput) &&
                uint.TryParse(newPrice, out uint NewPriceOutput) &&
                ushort.TryParse(newCtegory, out ushort newCategoryOutput))
            {
                hotel.numbers.Add(new Numbers(hotel.numbers.Max(c => c._id) + 1, NewFloorOutput, NewPlaceOutput, NewPriceOutput, newCategoryOutput));
                logger.Log($"Добавлена новая запись в таблицу номера с полями {newFloor}, {newPlace}, {newPrice}, {newCtegory}");
            }
            else
            {
                Console.WriteLine("Некорректные данные для номера.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при добавлении номера: {ex.Message}");
            logger.Log($"Ошибка при добавлении номера: {ex.Message}");
        }
    }

    private void AddReservation()
    {
        try
        {
            AddClients();
            AddNumber();


            Console.WriteLine("Введите новую дату бронирования (вида дд.мм.гггг):");
            string newReservationDate = Console.ReadLine();

            Console.WriteLine("Введите новую дату заезда:");
            string newReservationChekInDate = Console.ReadLine();

            Console.WriteLine("Введите новую дату выезда:");
            string newDepartureData = Console.ReadLine();

            if (!string.IsNullOrWhiteSpace(newReservationDate) &&
                !string.IsNullOrEmpty(newReservationChekInDate) &&
                !string.IsNullOrEmpty(newDepartureData))
            {
                hotel.reservation.Add(new Reservation(
                    hotel.reservation.Max(r => r.id) + 1,
                    hotel.clients.Max(c => c._id),
                    hotel.numbers.Max(c => c._id),
                    new CustomDate(newReservationDate),
                    new CustomDate(newReservationChekInDate),
                    new CustomDate(newDepartureData))
                );
                logger.Log($"Добавлена новая запись в таблицу Бронирование с полями {newReservationDate}, {newReservationChekInDate}, {newDepartureData}");
            }
            else
            {
                Console.WriteLine("Некорректные данные для бронирования.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при добавлении бронирования: {ex.Message}");
            logger.Log($"Ошибка при добавлении бронирования: {ex.Message}");
        }
    }

    // Запросы
    public void GetTotalClients()
    {
        try
        {
            Console.WriteLine($"Общее количество клиентов: {hotel.clients.Count()}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при получении количества клиентов: {ex.Message}");
            logger.Log($"Ошибка при получении количества клиентов: {ex.Message}");
        }
    }

    public void GetReservedNumbers()
    {
        try
        {
            var reservedNumbers = hotel.reservation
                .Select(r => hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._id)
                .Where(n => n != null)
                .ToList();

            Console.WriteLine("Зарезервированные номера: ");
            foreach (var reservedNumber in reservedNumbers)
            {
                Console.Write(reservedNumber + " ");
            }
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при получении зарезервированных номеров: {ex.Message}");
            logger.Log($"Ошибка при получении зарезервированных номеров: {ex.Message}");
        }
    }

    public void GetClientsWithReservation()
    {
        try
        {
            var clientsWithReservation = hotel.reservation
                .Select(r => new
                {
                    ClinetName = hotel.clients.FirstOrDefault(c => c._id == r.clientsId)?._secondName,
                    RoomsNumber = hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._id,
                    Price = hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._price
                })
                .Where(res => res.ClinetName != null && res.RoomsNumber != null)
                .Select(res => $"{res.ClinetName,-15}| {res.RoomsNumber,-10}| {res.Price,-5}|")
                .ToList();

            Console.WriteLine();
            Console.WriteLine("Клиент         | ID комнаты | Цена  |");
            foreach (var client in clientsWithReservation)
            {
                Console.WriteLine(client);
            }

            if (!clientsWithReservation.Any())
            {
                Console.WriteLine("Нет клиентов с бронированием.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при получении клиентов с бронированием: {ex.Message}");
            logger.Log($"Ошибка при получении клиентов с бронированием: {ex.Message}");
        }
    }

    public void GetRevenueForCity(string city) 
    {
        try
        {
            if (string.IsNullOrWhiteSpace(city))
            {
                Console.WriteLine("Город не может быть пустым.");
                return;
            }

            decimal totalCost = hotel.reservation
                .Where(r => hotel.clients.FirstOrDefault(c => c._id == r.clientsId)?._adres == city)
                .Sum(r => hotel.numbers.FirstOrDefault(n => n._id == r.numbersId)?._price ?? 0);

            Console.WriteLine($"Общая стоимость бронирования для города {city}: {totalCost}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при получении дохода для города: {ex.Message}");
            logger.Log($"Ошибка при получении дохода для города: {ex.Message}");
        }
    }

    public void SaveChagesInXls()
    {
        try
        {
            Workbook wb = new Workbook(PathToXlsFile);
            Worksheet wsClients = wb.Worksheets["Клиенты"];
            Worksheet wsReservation = wb.Worksheets["Бронирование"];
            Worksheet wsNumbers = wb.Worksheets["Номера"];

            // Сохранение клиентов
            for (int i = 0; i < hotel.clients.Count; i++)
            {
                var client = hotel.clients[i];
                wsClients.Cells[$"A{i + 2}"].PutValue(client._id);
                wsClients.Cells[$"B{i + 2}"].PutValue(client._firstName);
                wsClients.Cells[$"C{i + 2}"].PutValue(client._secondName);
                wsClients.Cells[$"D{i + 2}"].PutValue(client._patronymic);
                wsClients.Cells[$"E{i + 2}"].PutValue(client._adres);
            }

            // Сохранение номеров
            for (int i = 0; i < hotel.numbers.Count; i++)
            {
                var number = hotel.numbers[i];
                wsNumbers.Cells[$"A{i + 2}"].PutValue(number._id);
                wsNumbers.Cells[$"B{i + 2}"].PutValue(number._floor);
                wsNumbers.Cells[$"C{i + 2}"].PutValue(number._numbeerOfplaces);
                wsNumbers.Cells[$"D{i + 2}"].PutValue(number._price);
                wsNumbers.Cells[$"E{i + 2}"].PutValue(number._category);
            }

            // Сохранение бронирований
            for (int i = 0; i < hotel.reservation.Count; i++)
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
            logger.Log("Изменения успешно сохранены в файл.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при сохранении изменений в файл Excel: {ex.Message}");
            logger.Log($"Ошибка при сохранении изменений в файл Excel: {ex.Message}");
        }
    }
}