public struct CustomDate
{
    public int Day { get; }
    public int Month { get; }
    public int Year { get; }

    public CustomDate(string date)
    {
        // Разделяем строку даты по точке
        //var parts = date.Split(' ');

        
        var parts = date.Split(' ')[0].Split('.');
        if (parts.Length != 3)
            throw new FormatException("Неверный формат даты. Ожидался формат дд.мм.гггг.");


        // Преобразуем части в целые числа
        Day = int.Parse(parts[0]);
        Month = int.Parse(parts[1]);
        Year = int.Parse(parts[2]);

        if (Day < 1 || Day > 31)
            throw new ArgumentOutOfRangeException(nameof(Day), "День должен быть от 1 до 31.");
        if (Month < 1 || Month > 12)
            throw new ArgumentOutOfRangeException(nameof(Month), "Месяц должен быть от 1 до 12.");
        if (Year < 1)
            throw new ArgumentOutOfRangeException(nameof(Year), "Год должен быть положительным.");
    }

    public override string ToString()
    {
        return $"{Day}.{Month}.{Year}";
    }
}


public class Reservation
{
    public uint id { get; set; }
    public uint clientsId { get; set; }
    public uint numbersId { get; set; }
    public CustomDate ReservationData { get; set; }
    public CustomDate ChekInData { get; set; }
    public CustomDate departureDate { get; set; }

    public Reservation(uint ID, uint clientsID, uint NAMBID, CustomDate ResDate, CustomDate ChekDate, CustomDate DepDate)
    {
        id = ID;
        clientsId = clientsID;
        numbersId = NAMBID;
        ReservationData = ResDate;
        ChekInData = ChekDate;
        departureDate = DepDate;

    }

    public override string ToString()
    {
        return $"{id,-5} | {clientsId, -10} | {numbersId, -10} | {ReservationData, -15} | {ChekInData, -15} | {departureDate, -15}";
    }


}