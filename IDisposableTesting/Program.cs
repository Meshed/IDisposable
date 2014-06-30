namespace IDisposableTesting
{
    class Program
    {
        static void Main(string[] args)
        {
            StandardMethodTesting();
            UsingMethodTesting();
        }

        static private void StandardMethodTesting()
        {
            var wordService = new WordService();

            wordService.OpenDocument();
            wordService.CloseDocument();
            wordService.QuitApplication();
            wordService.Dispose();
        }

        static private void UsingMethodTesting()
        {
            using (var wordService = new WordService())
            {
                wordService.OpenDocument();
                wordService.CloseDocument();
                wordService.QuitApplication();
            }
        }
    }
}
