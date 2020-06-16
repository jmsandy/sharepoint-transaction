// Copyright 2020 Polimorfismo - José Mauro da Silva Sandy
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Controls the background tasks that are performed for the preparation of commands.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-25 11:27:26 PM</Date>
    internal class SharePointBackgroundTasks : IDisposable
    {
        #region Properties

        internal CancellationToken Token => CancellationTokenSource.Token;

        private ConcurrentBag<Task> Tasks { get; set; } = new ConcurrentBag<Task>();

        private CancellationTokenSource CancellationTokenSource { get; set; } = new CancellationTokenSource();

        #endregion

        #region Constructors / Finalizers

        public SharePointBackgroundTasks()
        {
        }

        ~SharePointBackgroundTasks() => Dispose(false);

        #endregion

        #region Methods

        public void Action(Action action)
        {
            Tasks.Add(Task.Run(action, Token));
        }

        public void Wait(int seconds)
        {
            _ = Task.WhenAny(Task.WhenAll(Tasks), Task.Delay(TimeSpan.FromSeconds(seconds))).Result;
        }

        public void Cancel()
        {
            CancellationTokenSource.Cancel();
        }

        public void Clear()
        {
            Tasks = new ConcurrentBag<Task>();
            CancellationTokenSource = new CancellationTokenSource();
        }

        public bool AllTasksCompletedSuccess()
        {
            return Tasks.All(task => task.Status == TaskStatus.RanToCompletion);
        }

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (disposing)
            {
                Cancel();
            }
        }

        #endregion
    }
}
