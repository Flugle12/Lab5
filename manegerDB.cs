using System.Net;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using static System.Runtime.InteropServices.JavaScript.JSType;

public class manegerBD
{
    private string PathToXlsFile;
    private BD hotel;

    public manegerBD(string path)
    {
        PathToXlsFile = path;
        hotel = new BD(PathToXlsFile);
    }

    public void OutputBD()
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
    }

    public void CorrectionElement(uint key)
    {
        Console.WriteLine("Введите номер таблицы корректировки: ");
        Console.WriteLine("1. Клиенты");
        Console.WriteLine("2. Номера");
        Console.WriteLine("3. Бронирования");
        string ch = Console.ReadLine();

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
                Console.WriteLine();
                break;
        }
    }

    public void AddElement()
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

    private void DeleteElementFromReservation(uint key)
    {
        var deleteReservation = hotel.reservation.FirstOrDefault(b => b.id == key);
        if (deleteReservation != null)
        {
            var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == deleteReservation.numbersId);
            var deleteClients = hotel.clients.FirstOrDefault(c => c._id == deleteReservation.clientsId);

            hotel.reservation.Remove(deleteReservation);
            hotel.numbers.Remove(deleteNumbers);
            hotel.clients.Remove(deleteClients);
        }
    }

    private void DeleteElementFromClients(uint clientsKey)
    {
        var deleteClients = hotel.clients.FirstOrDefault(c => c._id == clientsKey);
        if(deleteClients != null)
        {
            var deleteReservation = hotel.reservation.FirstOrDefault(r => r.clientsId == deleteClients._id);
            var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == deleteReservation.numbersId);

            hotel.reservation.Remove(deleteReservation);
            hotel.numbers.Remove(deleteNumbers);
            hotel.clients.Remove(deleteClients);
        }
    }

    private void DeleteElementFromNumbers(int numberKey)
    {
        var deleteNumbers = hotel.numbers.FirstOrDefault(n => n._id == numberKey);
        if( deleteNumbers != null)
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
        if( client == null )
        {
            return;
        }

        Console.WriteLine("Введите новую фамилию(чтобы не менять оставьте пустым)");
        string newSecName = Console.ReadLine();
        if (!string.IsNullOrEmpty(newSecName)) client._firstName = newSecName;

        Console.WriteLine("Введите новое имя");
        string newfirstName = Console.ReadLine();
        if(!string.IsNullOrEmpty(newfirstName)) client._secondName = newfirstName;

        Console.WriteLine("введите новое отчество");
        string newPatr = Console.ReadLine();
        if(!string.IsNullOrEmpty(newPatr)) client._patronymic = newPatr;

        Console.WriteLine("Введите новый адрес");
        string newAdress = Console.ReadLine();
        if (!string.IsNullOrEmpty(newAdress)) client._adres = newAdress;
    }

    private void CorrectionElementsNumbers(uint key)
    {
        var number = hotel.numbers.FirstOrDefault(n => n._id == key);
        if( number == null)
        {
            return;
        }

        Console.WriteLine("Введите новый этаж: ");
        string newFloor = Console.ReadLine();
        if(ushort.TryParse(newFloor, out ushort newFloorInput)) number._floor = newFloorInput;

        Console.WriteLine("Введите новое количество мест: ");
        string newPlace = Console.ReadLine();
        if(ushort.TryParse(newPlace, out ushort newPlacesInput)) number._numbeerOfplaces = newPlacesInput;

        Console.WriteLine("Введите новую цену: ");
        string newPrice = Console.ReadLine();
        if(uint.TryParse(newPrice, out uint newPriceInput)) number._price = newPriceInput;

        Console.WriteLine("ВВедите новую категорию номера: ");
        string newCtegory = Console.ReadLine();
        if(ushort.TryParse(newCtegory, out ushort newCtegoryInput)) number._category = newCtegoryInput;
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

    }

    private void AddClients()
    {
        Console.WriteLine("Введите новую фамилию(чтобы не менять оставьте пустым)");
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
    }

    private void AddReservation()
    {
        Console.WriteLine("Введите новую дату бронирования вида дд.мм.гггг");
        string newReservationDate = Console.ReadLine();

        Console.WriteLine("Введите новую дату заезда");
        string newReservationChekInDate = Console.ReadLine();

        Console.WriteLine("Введите новую дату выезда");
        string newDepartureData = Console.ReadLine();

        if(string.IsNullOrWhiteSpace(newReservationDate) && !string.IsNullOrEmpty(newReservationDate) && !string.IsNullOrEmpty(newDepartureData))
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
        AddClients();
        AddNumber();
    }

    public void GetRevenueForCity(string city)
    {
        decimal totalCost = hotel.reservation
            .Where(r => hotel.clients.FirstOrDefault(c => c._id == r.clientsId)?._adres == city).Sum(r => hotel.numbers.FirstOrDefault(n=>n._id == r.numbersId)?._price ?? 0);

        Console.WriteLine(totalCost);
    }

    public void GetReservationOnClients()
    {
        var reservationWithClients = from res in hotel.reservation
                                     join client in hotel.clients on res.clientsId equals client._id
                                     select new
                                     {
                                         res.id,
                                         res.ChekInData,
                                         ClientName = client._firstName + " " + client._secondName
                                     };

        foreach (var cl  in reservationWithClients)
        {
            Console.WriteLine(cl);
        }
    }


}