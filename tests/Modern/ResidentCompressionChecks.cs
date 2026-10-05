#nullable enable
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace SharePointExplorer.Tests
{
    internal static class ResidentCompressionChecks
    {
        public static void Run()
        {
            byte[] original=Encoding.UTF8.GetBytes("Original resident bytes remain present after external compression.");
            foreach(int compressedLength in new[] {0,1,original.Length,original.Length+8})
            {
                StoredChunk resident=Read(original,original.Length,compressedLength,true);
                Check(resident.Content.SequenceEqual(original),"External ABS compression metadata altered available original resident bytes.");
                using var output=new MemoryStream();new PlainDocumentDecoder().Write(new[] {resident},original.Length,output);
                Check(output.ToArray().SequenceEqual(original),"Resident native recovery attempted to decompress metadata-marked original bytes.");
            }
            // Actual compressed-looking bytes can also be original file content.
            // Only the declared exact resident length is authoritative here.
            byte[] gzipFile={31,139,8,0,0,0,0,0,0,3,3,0,0,0,0,0,0,0,0,0};
            Check(Read(gzipFile,gzipFile.Length,10,true).Content.SequenceEqual(gzipFile),"A resident gzip file was implicitly decompressed as external compression.");
            Reject<ContentUnavailableException>(()=>Read(null,999,20,true));
            Reject<ContentUnavailableException>(()=>Read(original,original.Length,20,true,rbs:new byte[] {1}));
            Reject<NotSupportedException>(()=>Read(original,original.Length,20,false));
            Reject<InvalidDataException>(()=>Read(original,-1,20,true));
            Reject<InvalidDataException>(()=>Read(original,original.Length,-1,true));
            Reject<InvalidDataException>(()=>Read(original,original.Length+1,20,true));
            Reject<InvalidDataException>(()=>Read(new byte[] {1},Int32.MaxValue,1,true));
            Reject<InvalidDataException>(()=>Read(gzipFile,1000000,gzipFile.Length,true));
            Check(Read(Array.Empty<byte>(),0,0,true).Content.Length==0,"An exact empty resident stream with external metadata was rejected.");
            Console.WriteLine("PASS resident inline bytes with external ABS compression metadata, no implicit decompression, RBS/unavailable and bounded size corruption guards");
        }
        private static StoredChunk Read(byte[]? content,int size,int? compressed,bool abs,byte[]? rbs=null)
        {
            using var table=new DataTable();
            Type[] types={typeof(long),typeof(long),typeof(byte),typeof(byte),typeof(int),typeof(byte[]),typeof(int),typeof(Guid),typeof(byte[])};
            for(int index=0;index<types.Length;index++)table.Columns.Add("c"+index,types[index]);
            table.Rows.Add(1L,1L,(byte)0,(byte)0,size,rbs==null?DBNull.Value:rbs,compressed.HasValue?compressed.Value:DBNull.Value,
                abs?Guid.NewGuid():DBNull.Value,content==null?DBNull.Value:content);
            using var reader=table.CreateDataReader();reader.Read();return SqlDocumentChunkStore.ReadChunk(reader);
        }
        private static void Check(bool condition,string message){if(!condition)throw new Exception(message);}
        private static void Reject<T>(Action action)where T:Exception {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name+".");}
    }
}