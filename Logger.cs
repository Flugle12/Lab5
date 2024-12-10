public class Logger
{
    private string _logFilePath;

    public Logger(string newPath)
    {
        _logFilePath = newPath;
        CreateLogFileIfNotExists();
    }

    public void SetNewFilePath(string newPath)
    {
        _logFilePath = newPath;
    }

    public void Log(string messege)
    {
        using(StreamWriter writer = new StreamWriter(_logFilePath, true))
        {
            writer.WriteLine($"{DateTime.Now}: {messege}");
        }
    }

    private void CreateLogFileIfNotExists()
    {
        if (!File.Exists(_logFilePath))
        {
            using (File.Create(_logFilePath)) // Создает файл, если он не существует
            {
                // Файл будет создан и закрыт
            }
        }
    }
}