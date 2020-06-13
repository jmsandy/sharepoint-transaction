using System;
using Serilog;
using Polimorfismo.SharePoint.Transaction.Logging;

namespace Polimorfismo.SharePointOnline.Transaction.Tests.Logging
{
    /// <summary>
    /// Implementation of the log interface using Serilog.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-12 08:07:26 PM</Date>
    internal class SharePointSerilog : ISharePointLogging
    {
        private readonly ILogger Logger;

        public SharePointSerilog()
        {
            Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File("sharepoint-log.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
        }

        public void Debug(string message, params object[] objs)
        {
            Logger.Debug(message, objs);
        }

        public void Info(string message, params object[] objs)
        {
            Logger.Information(message, objs);
        }

        public void Warn(string message, params object[] objs)
        {
            Logger.Warning(message, objs);
        }

        public void Warn(Exception exception, string message, params object[] objs)
        {
            Logger.Warning(exception, message, objs);
        }

        public void Error(Exception exception, string message, params object[] objs)
        {
            Logger.Error(exception, message, objs);
        }

        public void Error(string message, params object[] objs)
        {
            Logger.Error(message, objs);
        }
    }
}
