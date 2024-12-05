using System.Runtime.CompilerServices;

public class Сlients
{
    public uint _id { get; set; }
    public string? _firstName { get; set; }
    public string? _secondName { get; set; }
    public string? _patronymic { get; set; }
    public string? _adres { get; set; }

    public override string ToString()
    {
        return $"{_id,-3} | {_firstName, -20} | {_secondName, -10} | {_patronymic, -15} | {_adres,-15}";
    }

    public Сlients(uint id, string FN, string SN, string patr, string adres)
    {
        _id = id ;
        _firstName = FN ;
        _secondName = SN ;
        _patronymic = patr ;
        _adres = adres ;
    }

}