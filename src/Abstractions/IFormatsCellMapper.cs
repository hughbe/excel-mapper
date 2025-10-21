using System;

namespace ExcelMapper.Abstractions;

internal interface IFormatsCellMapper : ICellMapper
{
    string[] Formats { get; set; }

    IFormatProvider? Provider { get; set; }
}
