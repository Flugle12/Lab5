public class Numbers
{
    public uint _id { get; set; }
    public ushort _floor { get; set; }
    public ushort _numbeerOfplaces { get; set; }
    public uint _price {  get; set; }
    public ushort _category { get; set; }

    public Numbers(uint id, ushort floor, ushort numb, uint price, ushort category)
    {
        _id = id;
        _floor = floor;
        _numbeerOfplaces = numb;
        _price = price;
        _category = category;
    }

    public override string ToString()
    {
        return $"{_id,-5} | {_floor, -3} | {_numbeerOfplaces, -3} | {_price, -8} | {_category,-3}";
    }
}
