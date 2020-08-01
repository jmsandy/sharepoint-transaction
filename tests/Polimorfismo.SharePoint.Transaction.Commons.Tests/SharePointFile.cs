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

using System.IO;

namespace Polimorfismo.SharePoint.Transaction.Commons.Tests
{
    /// <summary>
    /// Represents a document in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 09:10:18 AM</Date>
    public class SharePointFile : ISharePointFile
    {
        public string ListName => "DocumentsList";

        public int Id { get; set; }

        [SharePointField("Description0")]
        public string Description { get; set; }

        public string Folder { get; set; }

        public string FileName { get; set; }

        public Stream InputStream { get; set; }
    }
}
