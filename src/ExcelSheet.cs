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

        internal ExcelSheet(IExcelDataReader reader, int index, ExcelImporter importer)
        {
            Reader = reader;
            Name = reader.Name;
            if (reader.VisibleState == "visible")
            {
                Visibility = ExcelSheetVisibility.Visible;
            }
            else if (reader.VisibleState == "hidden")
            {
                Visibility = ExcelSheetVisibility.Hidden;
            }
            else
            {
                Visibility = ExcelSheetVisibility.VeryHidden;
            }
            Index = index;
            Importer = importer;
        }

        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the visibility of the sheet.
        /// </summary>
        public ExcelSheetVisibility Visibility { get; }

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

        /// <summary>
        /// Gets the index of the row currently being mapped.
        /// </summary>
        public int CurrentRowIndex { get; private set; } = -1;

        private ExcelImporter Importer { get; }

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

            Importer.MoveToSheet(this);
            ReadPastHeading();

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
        /// Maps each row within the range specified to an object using a registered mapping. If no map is registered for this
        /// type then the type will be automapped. This method will not read the sheet's heading.
        /// </summary>
        /// <param name="startIndex">The zero-based index from the first row of the document (including the header) of the range of rows to map from.</param>
        /// <param name="count">The number of rows to read and map.</param>
        /// <typeparam name="T">The type of the object to map each row to.</typeparam>
        /// <returns>A list of objects of type T mapped from each row within the range specified.</returns>
        public IEnumerable<T> ReadRows<T>(int startIndex, int count)
        {
            if (startIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startIndex), startIndex, "Start index cannot be negative.");
            }
            if (count < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(count), count, "The number of rows cannot be negative.");
            }

            CurrentRowIndex = startIndex;
            for (int i = 0; i < count; i++)
            {
                yield return ReadRow<T>();
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
            Importer.MoveToSheet(this);

            value = default(T);
            if (!Reader.Read())
            {
                return false;
            }

            CurrentRowIndex++;

            if (Importer.Configuration.SkipBlankLines)
            {
                bool RowEmpty()
                {
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        if (!(Reader.GetValue(i) is null))
                        {
                            return false;
                        }
                    }

                    return true;
                }

                while (RowEmpty())
                {
                    if (!Reader.Read())
                    {
                        return false;
                    }

                    CurrentRowIndex++;
                }
            }

            if (!Importer.Configuration.TryGetClassMap<T>(out IMap classMap))
            {
                if (!HasHeading)
                {
                    throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\" as the sheet has no heading.");
                }

                if (!AutoMapper.TryAutoMap<T>(FallbackStrategy.ThrowIfPrimitive, out IMap map))
                {
                    throw new ExcelMappingException($"Cannot auto-map type \"{typeof(T)}\".");
                }

                classMap = map;
                Importer.Configuration.RegisterClassMap(typeof(T), classMap);
            }

            bool result = classMap.TryGetValue(this, CurrentRowIndex, Reader, null, out object valueObject);
            value = (T)valueObject;
            return result;
        }

        internal void ReadPastHeading()
        {
            for (int i = 0; i <= HeadingIndex; i++)
            {
                if (!Reader.Read())
                {
                    throw new ExcelMappingException($"Sheet \"{Name}\" has no heading.");
                }
            }
        }
    }
}
