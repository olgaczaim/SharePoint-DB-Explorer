using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using SharePointExplorer;

namespace SharePointExplorer.Tests
{
    internal static class CatalogChecks
    {
        internal static void Run()
        {
            if((int)NodeKind.Site!=0 || (int)NodeKind.Library!=1 || (int)NodeKind.Folder!=2 || (int)NodeKind.File!=3)
                throw new Exception("The existing catalog node-kind values changed.");
            if(!new Node { Kind=NodeKind.List }.IsContainer || new Node { Kind=NodeKind.ListItem }.IsContainer || new Node { Kind=NodeKind.File }.IsContainer)
                throw new Exception("List navigation or browse-only item semantics are incorrect.");
            TitleField title=ListFieldSchema.ReadTitle(Store("16.0.0.14326.0.0<FieldRef Name=\"Title\" ColName=\"nvarchar7\" RowOrdinal=\"2\"/>"));
            if(title==null || title.Column!="nvarchar7" || title.RowOrdinal!=2) throw new Exception("The catalog guessed title storage instead of reading its declared column and ordinal.");
            title=ListFieldSchema.ReadTitle(Store("16.0.0.14326.0.0<Field Name=\"Title\" ColName=\"ntext2\" RowOrdinal=\"0\"/><Field Name=\"LinkTitle\"><FieldRefs><FieldRef Name=\"Title\"/></FieldRefs></Field>"));
            if(title==null || title.Column!="ntext2" || title.RowOrdinal!=0) throw new Exception("A nested display reference replaced the title storage mapping.");
            if(ListFieldSchema.ReadTitle(Store("16.0.0.14326.0.0<FieldRef Name=\"OtherField\" ColName=\"nvarchar1\"/>"))!=null)
                throw new Exception("An arbitrary first string field became the item's title.");
            MustReject(Store("16.0.0.14326.0.0<FieldRef Name=\"Title\" ColName=\"nvarchar1]; SELECT 1--\"/>"),"unsafe column name");
            MustReject(Store("16.0.0.14326.0.0<FieldRef Name=\"Title\" ColName=\"nvarchar1\" RowOrdinal=\"256\"/>"),"invalid row ordinal");
            MustReject(Store("16.0.0.14326.0.0<FieldRef Name=\"Title\" ColName=\"nvarchar1\"/><FieldRef Name=\"Title\" ColName=\"nvarchar2\"/>"),"conflicting title mappings");
            byte[] corrupt=Store("16.0.0.14326.0.0<FieldRef Name=\"Title\" ColName=\"nvarchar1\"/>");
            corrupt[corrupt.Length-1]^=1;
            MustReject(corrupt,"corrupted compressed-metadata checksum");
            byte[] incorrectLength=Store("16.0.0.14326.0.0<FieldRef Name=\"Title\" ColName=\"nvarchar1\"/>");
            incorrectLength[8]--;
            MustReject(incorrectLength,"incorrect compressed-metadata length");
        }
        private static void MustReject(byte[] value,string reason)
        {
            try { ListFieldSchema.ReadTitle(value); }
            catch(InvalidDataException) { return; }
            catch(NotSupportedException) { return; }
            throw new Exception("The list metadata parser accepted "+reason+".");
        }
        private static byte[] Store(string text)
        {
            byte[] original=Encoding.UTF8.GetBytes(text);
            using(var result=new MemoryStream())
            {
                result.Write(new byte[] { 0xA8,0xA9,0x30,0x31,12,0,0,0 },0,8);
                byte[] length=BitConverter.GetBytes(original.Length);result.Write(length,0,4);
                result.WriteByte(0x78);result.WriteByte(0x9C);
                using(var compressor=new DeflateStream(result,CompressionMode.Compress,true)) compressor.Write(original,0,original.Length);
                uint sum1=1,sum2=0;
                foreach(byte value in original) { sum1=(sum1+value)%65521;sum2=(sum2+sum1)%65521; }
                uint checksum=(sum2<<16)|sum1;
                result.WriteByte((byte)(checksum>>24));result.WriteByte((byte)(checksum>>16));result.WriteByte((byte)(checksum>>8));result.WriteByte((byte)checksum);
                return result.ToArray();
            }
        }
    }
}