using System.IO;
using System.Text;

namespace ExcelMapper.Tests
{
    public static class Helpers
    {
        public static bool Initialized { get; private set; }

        public static ExcelImporter GetImporter(string name) => new ExcelImporter(GetResource(name));

        public static Stream GetResource(string name)
        {
            if (!Initialized)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                Initialized = true;
            }

            string filePath = Path.GetFullPath(Path.Combine("Resources", name));
            return File.OpenRead(filePath);
        }
    }
}
