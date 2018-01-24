using System;
using System.Collections.Generic;
using ExcelDataReader;
using ExcelMapper.Utilities;

namespace ExcelMapper
{
    /// <summary>
    /// An object that represents a single sheet of an excel document.
    /// </summary>
    public class ExcelSheet
    {
        private bool _hasHeading = true;
        private int _headingIndex = 0;

        internal ExcelSheet(IExcelDataReader reader, int index, ExcelImporterConfiguration configuration)
        {
            Reader = reader;
            Name = reader.Name;
            Index = index;
            Configuration = configuration;
        }

        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the zero-based index of the sheet where 0 is the first sheet in the document.
        /// </summary>
        public int Index { get; }

        /// <summary>
        /// Gets or sets whether the sheet has a heading. This is true by default.
        /// </summary>
        public bool HasHeading
        {
            get => _hasHeading;
            set
            {
                if (Heading != null)
                {
                    throw new InvalidOperationException("The heading has already been read. Set this property before reading any rows.");
                }

                _hasHeading = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based index of row containing the heading. This is 0 (the first row) by default.
        /// If the value is non-zero, all rows preceding the heading are skipped and not mapped.
        /// </summary>
        public int HeadingIndex
        {
            get => _headingIndex;
            set
            {
                if (value < 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(value), value, "The index of the heading must be positive or zero.");
                }
                if (!HasHeading)
                {
                    throw new InvalidOperationException("The sheet has no heading.");
                }
                if (Heading != null)
                {
                    throw new InvalidOperationException("The heading has already been read.");
                }

                _headingIndex = value;
            }
        }

        /// <summary>
        /// Gets the heading that was read from the sheet. This will return null if HasHeading is false
        /// or the heading has not been read yet by calling ReadHeading or ReadRows.
        /// </summary>
        public ExcelHeading Heading { get; private set; }

        private int CurrentIndex { get; set; } = -1;
        private ExcelImporterConfiguration Configuration { get; }
        private IExcelDataReader Reader { get; }

        /// <summary>
        /// Reads the heading of the sheet including column names and indices.
        /// </summary>
        /// <returns>An object that represents the heading of the sheet.</returns>
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

            for (int i = 0; i <= HeadingIndex; i++)
            {
                if (!Reader.Read())
                {
                    throw new ExcelMappingException($"Sheet \"{Name}\" has no heading.");
                }
            }

            var heading = new ExcelHeading(Reader);
            Heading = heading;
            return heading;
        }

        /// <summary>
        /// Maps each row of the sheet to an object using a registered mapping. If no map is registered for this
        /// type then the type will be automapped. This method will read the sheet's heading if the sheet has
        /// a heading and the heading has not yet been read.
        /// </summary>
        /// <typeparam name="T">The type of the object to map each row to.</typeparam>
        /// <returns>A list of objects of type T mapped from each row in the sheet.</returns>
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

        /// <summary>
        /// Maps a single row of a sheet to an object using a registered mapping. If no map is registered for this
        /// type then the type will be automapped. This method will not read the sheet's heading if the sheet has a
        /// heading and the heading has not yet been read. This method will throw if mapping fails or there are
        /// no more rows left.
        /// </summary>
        /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
        /// <returns>An object of type T mapped from a single row in the sheet.</returns>
        public T ReadRow<T>()
        {
            if (!TryReadRow(out T value))
            {
                throw new ExcelMappingException($"No more rows in \"{Name}\".");
            }

            return value;
        }

        /// <summary>
        /// Maps a single row of a sheet to an object using a registered mapping. If no map is registered for this
        /// type then the type will be automapped. This method will not read the sheet's heading if the sheet has a
        /// heading and the heading has not yet been read.
        /// </summary>
        /// <typeparam name="T">The type of the object to map a single row to.</typeparam>
        /// <param name="value">An object of type T mapped from a single row in the sheet.</param>
        /// <returns>False if there are no more rows in the sheet or the row cannot be mapped to an object, else false.</returns>
        public bool TryReadRow<T>(out T value)
        {
            value = default(T);
            if (!Reader.Read())
            {
                return false;
            }

            CurrentIndex++;

            if (!Configuration.TryGetClassMap<T>(out ExcelClassMap classMap))
            {
                if (!HasHeading)
                {
                    throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\" as the sheet has no heading.");
                }

                if (!AutoMapper.AutoMapClass(FallbackStrategy.ThrowIfPrimitive, out ExcelClassMap<T> autoClassMap))
                {
                    throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\".");
                }

                classMap = autoClassMap;
                Configuration.RegisterClassMap(autoClassMap);
            }

            value = (T)classMap.Execute(this, CurrentIndex, Reader);
            return true;
        }
    }
}
