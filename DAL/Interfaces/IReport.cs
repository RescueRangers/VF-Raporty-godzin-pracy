using System.Collections.Generic;

namespace DAL.Interfaces
{
    public interface IReport
    {
        List<Employee> Employees { get; set; }
        List<Header> Headers { get; }
        List<Translation> NotTranslatedHeaders { get; }

        bool AreHeadersTranslated { get;  }
    }
}