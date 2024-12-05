using Aspose.Cells;

public class BD
{
    public List<Reservation>? reservation;
    public List<Numbers>? numbers;
    public List<Сlients>? clients;

    public BD(string PathToXlsFile)
    {
        reservation = new List<Reservation>();
        numbers = new List<Numbers>();
        clients = new List<Сlients>();

        Workbook workbook = new Workbook(PathToXlsFile);
        WorksheetCollection worksheets = workbook.Worksheets;

        Worksheet clientsSheet = worksheets[0];


        for (int i = 1; i <= clientsSheet.Cells.MaxRow; i++)
        {
            var cellValue = clientsSheet.Cells[i, 0].Value;
            if (cellValue != null && int.TryParse(cellValue.ToString(), out int doubleValue))
            {
                uint id = Convert.ToUInt32(doubleValue);
                string? firstname = clientsSheet.Cells[i, 1].Value?.ToString();
                string? secondName = clientsSheet.Cells[i, 2].Value?.ToString();
                string? patronymic = clientsSheet.Cells[i, 3].Value?.ToString();
                string? adress = clientsSheet.Cells[i, 4].Value?.ToString();

                clients.Add(new Сlients(id, firstname, secondName, patronymic, adress));
            }
        }

        Worksheet reservationSheet = worksheets[1];

        for (int i = 1; i <= reservationSheet.Cells.MaxRow; i++)
        {
            var cellValue = reservationSheet.Cells[i, 0].Value;
            if (cellValue != null && double.TryParse(cellValue.ToString(), out double doubleValue))
            {
                uint id = Convert.ToUInt32(doubleValue);
                uint clientsId = Convert.ToUInt32(reservationSheet.Cells[i, 1].Value);
                uint numbersId = Convert.ToUInt32(reservationSheet.Cells[i, 2].Value);
                //string reservationDate = reservationSheet.Cells[i, 3].Value.ToString();
                //Console.WriteLine(reservationDate);
                CustomDate reservationDate = new CustomDate(reservationSheet.Cells[i, 3].Value.ToString());
                CustomDate ChekInDate = new CustomDate(reservationSheet.Cells[i, 4].Value.ToString());
                CustomDate departureDate = new CustomDate(reservationSheet.Cells[i, 5].Value.ToString());

                reservation.Add(new Reservation(id, clientsId, numbersId, reservationDate, ChekInDate, departureDate));
                //Console.WriteLine(reservationDate);
            }
        }

        Worksheet numbersSheet = worksheets[2];

        for (int i = 1; i <= numbersSheet.Cells.MaxRow; i++)
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
    }


}
