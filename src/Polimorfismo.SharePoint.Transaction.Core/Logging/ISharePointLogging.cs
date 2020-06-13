using System;

namespace Polimorfismo.SharePoint.Transaction.Logging
{
    /// <summary>
    /// Interface to trigger the library logging process.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-12 07:47:43 PM</Date>
    public interface ISharePointLogging
    {
        void Debug(string message, params object[] objs);

        void Info(string message, params object[] objs);

        void Warn(string message, params object[] objs);

        void Warn(Exception exception, string message, params object[] objs);

        void Error(string message, params object[] objs);

        void Error(Exception exception, string message, params object[] objs);
    }
}
