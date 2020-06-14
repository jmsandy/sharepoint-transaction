using System.IO;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Interface representing an document in the SharePoint library. 
    /// It can be having a representation for each operation to be performed.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-14 08:46:57 AM</Date>
    public interface ISharePointFile : ISharePointMetadata
    {
        string Folder { get; set; }

        string FileName { get; set; }

        Stream InputStream { get; set; }
    }
}
