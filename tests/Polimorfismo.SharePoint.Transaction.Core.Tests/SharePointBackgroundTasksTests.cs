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

using Xunit;
using System;
using FluentAssertions;
using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Core.Tests
{
    /// <summary>
    /// Tests for background tasks.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-15 09:08:42 PM</Date>
    public class SharePointBackgroundTasksTests
    {
        [Trait("Category", "SharePointCore - BackgroundTasks")]
        [Fact(DisplayName = "Cancel background tasks")]
        public void SharePointBackgroundTasks_Cancel_TasksNotCompleted()
        {
            // Arrange
            var backgroundTasks = new SharePointBackgroundTasks();

            // Act
            backgroundTasks.Action(() =>
            {
                Task.Delay(TimeSpan.FromSeconds(20));

                backgroundTasks.Token.ThrowIfCancellationRequested();
            });
            backgroundTasks.Cancel();

            // Assert
            backgroundTasks.AllTasksCompletedSuccess().Should().BeFalse();

            backgroundTasks.Dispose();
        }

        [Trait("Category", "SharePointCore - BackgroundTasks")]
        [Fact(DisplayName = "Background tasks completed")]
        public void SharePointBackgroundTasks_AllTasksCompletedSuccess_CompletedWithSuccess()
        {
            // Arrange
            var backgroundTasks = new SharePointBackgroundTasks();

            // Act
            backgroundTasks.Action(() =>
            {
                Task.Delay(TimeSpan.FromSeconds(20));

                backgroundTasks.Token.ThrowIfCancellationRequested();
            });
            backgroundTasks.Wait(30);

            // Assert
            backgroundTasks.AllTasksCompletedSuccess().Should().BeTrue();

            backgroundTasks.Dispose();
        }
    }
}
