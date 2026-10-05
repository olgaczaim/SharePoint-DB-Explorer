using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace SharePointExplorer
{
    // This is a retained metadata recovery artifact. It intentionally does not
    // claim to be a SharePoint deployment package or synthesize document bytes.
    public sealed class DeletedItemExportService
    {
        public ExportResult Export(DeletedListItemSnapshot snapshot,string directory)
        {
            if(snapshot==null)throw new ArgumentNullException("snapshot");
            Node item=snapshot.Item;DeletedIdentity.Validate(item,false);
            directory=System.IO.Path.GetFullPath(directory);Directory.CreateDirectory(directory);
            string name=DocumentExporter.SafeFileName((String.IsNullOrWhiteSpace(item.Name)?"Item "+item.ListItemId:item.Name)+" (deleted metadata).xml");
            string path=System.IO.Path.Combine(directory,name),stem=System.IO.Path.GetFileNameWithoutExtension(name);
            for(int suffix=2;File.Exists(path) || Directory.Exists(path);suffix++)path=System.IO.Path.Combine(directory,stem+" ("+suffix.ToString(CultureInfo.InvariantCulture)+").xml");
            string temporary=System.IO.Path.Combine(directory,"."+Guid.NewGuid().ToString("N")+".partial");
            try
            {
                using(var output=new FileStream(temporary,FileMode.CreateNew,FileAccess.Write,FileShare.None))
                {
                    using(var writer=XmlWriter.Create(output,new XmlWriterSettings {Encoding=new UTF8Encoding(false),Indent=true,CloseOutput=false}))
                    {
                        writer.WriteStartDocument();writer.WriteStartElement("DeletedListItemRecovery");writer.WriteAttributeString("FormatVersion","1");
                        writer.WriteAttributeString("SiteId",item.SiteId.ToString("D"));writer.WriteAttributeString("WebId",item.WebId.ToString("D"));
                        writer.WriteAttributeString("ListId",item.ListId.ToString("D"));writer.WriteAttributeString("DocumentId",item.Id.ToString("D"));
                        writer.WriteAttributeString("ItemId",item.ListItemId.Value.ToString(CultureInfo.InvariantCulture));writer.WriteAttributeString("ItemUniqueId",item.ItemUniqueId.Value.ToString("D"));
                        writer.WriteAttributeString("DeletionTransactionId",item.DeletionTransactionId);writer.WriteAttributeString("Path",item.Path ?? String.Empty);
                        writer.WriteAttributeString("UIVersion",item.UiVersion.ToString(CultureInfo.InvariantCulture));writer.WriteAttributeString("Level",item.Level.ToString(CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("InternalVersion",item.InternalVersion.ToString(CultureInfo.InvariantCulture));
                        if(item.DeletedAt.HasValue)writer.WriteAttributeString("DeletedAt",item.DeletedAt.Value.ToString("O",CultureInfo.InvariantCulture));
                        writer.WriteElementString("StoredFieldSchemaXml",snapshot.StoredFieldsXml);
                        writer.WriteStartElement("RetainedStorageValues");
                        foreach(DeletedListItemValue value in snapshot.Values)
                        {
                            writer.WriteStartElement("Value");writer.WriteAttributeString("RowOrdinal",value.RowOrdinal.ToString(CultureInfo.InvariantCulture));
                            writer.WriteAttributeString("Column",value.Column);writer.WriteAttributeString("SqlType",value.ValueType);
                            writer.WriteAttributeString("IsNull",value.IsNull?"true":"false");if(!value.IsNull)writer.WriteString(value.Value);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();writer.WriteEndElement();writer.WriteEndDocument();
                    }
                    output.Flush(true);
                }
                long length;string hash;
                using(FileStream input=File.OpenRead(temporary))using(SHA256 sha=SHA256.Create())
                {length=input.Length;hash=BitConverter.ToString(sha.ComputeHash(input)).Replace("-",String.Empty).ToLowerInvariant();}
                File.Move(temporary,path);
                return new ExportResult {Path=path,Bytes=length,Sha256=hash,Decoder="Retained deleted list-item metadata XML"};
            }
            finally {if(File.Exists(temporary))File.Delete(temporary);}
        }
    }
}