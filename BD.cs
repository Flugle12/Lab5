using Aspose.Cells;
using System.Collections.Generic;

public class BD
{
    private Logger logger;
    public List<Reservation>? reservation;
    public List<Numbers>? numbers;
    public List<Сlients>? clients;

    public BD(string PathToXlsFile, Logger logger)
    {
        this.logger = logger;
        reservation = new List<Reservation>();
        numbers = new List<Numbers>();
        clients = new List<Сlients>();

        try
        {
            Workbook workbook = new Workbook(PathToXlsFile);
            WorksheetCollection worksheets = workbook.Worksheets;

            // Загрузка клиентов
            LoadClients(worksheets[0]);

            // Загрузка бронирований
            LoadReservations(worksheets[1]);

            // Загрузка номеров
            LoadNumbers(worksheets[2]);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при загрузке данных из файла Excel: {ex.Message}");
        }
    }

    private void LoadClients(Worksheet clientsSheet)
    {
        for (int i = 1; i <= clientsSheet.Cells.MaxRow; i++)
        {
            try
            {
                var cellValue = clientsSheet.Cells[i, 0].Value;
                if (cellValue != null && int.TryParse(cellValue.ToString(), out int doubleValue))
                {
                    uint id = Convert.ToUInt32(doubleValue);
                    string? firstname = clientsSheet.Cells[i, 1].Value?.ToString();
                    string? secondName = clientsSheet.Cells[i, 2].Value?.ToString();
                    string? patronymic = clientsSheet.Cells[i, 3].Value?.ToString();
                    string? adress = clientsSheet.Cells[i, 4].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(firstname) && !string.IsNullOrWhiteSpace(secondName))
                    {
                        clients.Add(new Сlients(id, firstname, secondName, patronymic, adress));
                    }
                    else
                    {
                        Console.WriteLine($"Пропущена строка {i + 1}: имя или фамилия не могут быть пустыми.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке клиента на строке {i + 1}: {ex.Message}");
            }
        }
        logger.Log("Данные о клиентах успешно загружены");
    }

    private void LoadReservations(Worksheet reservationSheet)
    {
        for (int i = 1; i <= reservationSheet.Cells.MaxRow; i++)
        {
            try
            {
                var cellValue = reservationSheet.Cells[i, 0].Value;
                if (cellValue != null && double.TryParse(cellValue.ToString(), out double doubleValue))
                {
                    uint id = Convert.ToUInt32(doubleValue);
                    uint clientsId = Convert.ToUInt32(reservationSheet.Cells[i, 1].Value);
                    uint numbersId = Convert.ToUInt32(reservationSheet.Cells[i, 2].Value);

                    CustomDate reservationDate = new CustomDate(reservationSheet.Cells[i, 3].Value?.ToString());
                    CustomDate ChekInDate = new CustomDate(reservationSheet.Cells[i, 4].Value?.ToString());
                    CustomDate departureDate = new CustomDate(reservationSheet.Cells[i, 5].Value?.ToString());

                    reservation.Add(new Reservation(id, clientsId, numbersId, reservationDate, ChekInDate, departureDate));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке бронирования на строке {i + 1}: {ex.Message}");
            }
        }
        logger.Log("Данные о бронировании номеров успешно загружены");
    }

    private void LoadNumbers(Worksheet numbersSheet)
    {
        for (int i = 1; i <= numbersSheet.Cells.MaxRow; i++)
        {
            try
            {
                var cellValue = numbersSheet.Cells[i, 0].Value;
                if (cellValue != null && double.TryParse(cellValue.ToString(), out double doubleValue))
                {
                    uint id = Convert.ToUInt32(doubleValue);
                    ushort floor = Convert.ToUInt16(numbersSheet.Cells[i, 1].Value);
                    ushort numberOfPlace = Convert.ToUInt16(numbersSheet.Cells[i, 2].Value);
                    uint price = Convert.ToUInt32(numbersSheet.Cells[i, 3].Value);
                    ushort category = Convert.ToUInt16(numbersSheet.Cells[i, 4].Value);

                    numbers.Add(new Numbers(id, floor, numberOfPlace, price, category));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке номера на строке {i + 1}: {ex.Message}");
            }
        }
        logger.Log("Данные о номерах успешно загружены");
    }
}
