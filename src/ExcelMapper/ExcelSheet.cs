using System.Collections.Generic;
using ExcelDataReader;

namespace ExcelMapper
{
    public class ExcelSheet
    {
        internal ExcelSheet(IExcelDataReader reader, int index, ExcelImporterConfiguration configuration)
        {
            Reader = reader;
            Name = reader.Name;
            Index = index;
            Configuration = configuration;
            HasHeading = configuration.HasHeading == null ? true : configuration.HasHeading(this);
        }

        public string Name { get; }
        public int Index { get; }
        public bool HasHeading { get; }
        public ExcelHeading Heading { get; private set; }

        private int CurrentIndex { get; set; } = -1;
        private ExcelImporterConfiguration Configuration { get; }
        private IExcelDataReader Reader { get; }

        public ExcelHeading ReadHeading()
        {
            if (!HasHeading)
            {
                throw new ExcelMappingException($"Sheet \"{Name}\" has no heading.");
            }

            if (Heading != null)
            {
                throw new ExcelMappingException($"Already read heading in sheet \"{Name}\".");
            }

            Reader.Read();

            var heading = new ExcelHeading(Reader);
            Heading = heading;
            return heading;
        }

        public IEnumerable<T> ReadRows<T>()
        {
            if (HasHeading && Heading == null)
            {
                ReadHeading();
            }

            while (TryReadRow(out T row))
            {
                yield return row;
            }
        }

        public T ReadRow<T>()
        {
            if (!TryReadRow(out T value))
            {
                throw new ExcelMappingException($"No more rows in \"{Name}\".");
            }

            return value;
        }

        public bool TryReadRow<T>(out T value)
        {
            value = default(T);
            if (!Reader.Read())
            {
                return false;
            }

            CurrentIndex++;

            ExcelClassMap mapping = Configuration.GetMapping<T>();
            value = (T)mapping.Execute(this, CurrentIndex, Reader);
            return true;
        }
    }
}
