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
using System.Runtime.Serialization;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Exception for the communication process with Sharepoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-10 08:29:23 PM</Date>
    public class SharePointException : Exception, ISerializable
    {
        public object SharePointData { get; }

        public SharePointErrorCode ErrorCode { get; }

        public SharePointException(SharePointErrorCode errorCode, string message)
            : this(errorCode, null, message, null)
        {
        }

        public SharePointException(SharePointErrorCode errorCode, string message, Exception innerException)
            : this(errorCode, null, message, innerException)
        {
        }

        public SharePointException(SharePointErrorCode errorCode, object sharePointData, string message, Exception innerException)
            : base(message, innerException)
        {
            ErrorCode = errorCode;
            SharePointData = sharePointData;
        }
    }
}
