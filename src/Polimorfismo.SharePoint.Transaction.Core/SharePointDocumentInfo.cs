using System;
using System.IO;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Resources;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Set of information and contents of a document present in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-13 10:35:54 AM</Date>
    public class SharePointDocumentInfo
    {
        public int Id { get; }

        public string Name { get; }

        public byte[] Content { get; }

        public string ContentBase64 => Content != null ? Convert.ToBase64String(Content) : null;

        public int Length => IsFile && Content != null ? Content.Length : 0;

        public string Extension => IsFile && string.IsNullOrEmpty(Name) ? null : Path.GetExtension(Name);

        public bool IsFile { get; }

        public List<SharePointDocumentInfo> Documents { get; }

        public SharePointDocumentInfo(int id, string name, byte[] content, bool isFile)
        {
            Id = id;
            Name = name;
            IsFile = isFile;
            Content = content;
            Documents = new List<SharePointDocumentInfo>();
        }

        public void AddDocument(SharePointDocumentInfo document)
        {
            Documents.Add(document);
        }

        public void AddDocuments(IEnumerable<SharePointDocumentInfo> documents)
        {
            if (IsFile) Documents.AddRange(documents);

            throw new SharePointException(SharePointErrorCode.OnlyFoldersCanReceiveDocuments, SharePointMessages.ERR402);
        }
    }
}
