
public class Program
{
    public static void Main()
    {
        manegerBD meneger = new manegerBD("C:\\Users\\User\\Desktop\\LR5 Excel Files\\LR5-var9.xls");

        //meneger.OutputBD();

        //meneger.DeleteElementFromReservation(97560);
        //meneger.CorrectionElement(123);
        //meneger.OutputBD();
        meneger.GetRevenueForCity("г. Тында");
        meneger.GetReservationOnClients();
    }

}
