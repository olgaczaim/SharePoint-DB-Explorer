using System;

namespace SharePointExplorer
{
    // The session owns one catalog/storage lifetime. Recovery and file export
    // share the same engine, while the console only handles navigation/input.
    public sealed class RecoverySession : IDisposable
    {
        private readonly ISharePointCatalog catalog;
        private readonly IDocumentChunkStore store;
        private readonly RecoveryEngine engine;
        private readonly DocumentExporter exporter;
        private bool disposed;
        public RecoverySession(ISharePointCatalog catalog,IDocumentChunkStore store,DocumentDecoderRegistry registry)
        {
            if(catalog==null) throw new ArgumentNullException("catalog");
            if(store==null) throw new ArgumentNullException("store");
            if(registry==null) throw new ArgumentNullException("registry");
            this.catalog=catalog; this.store=store;
            engine=new RecoveryEngine(catalog,store,registry);
            exporter=new DocumentExporter(engine);
        }
        public static RecoverySession OpenSql(string server,string database)
        {
            return OpenSql(new SqlConnectionOptions { Server=server,Database=database });
        }
        public static RecoverySession OpenSql(SqlConnectionOptions options)
        {
            var catalog=new SqlRepository(options);
            catalog.ValidateSchema();
            return new RecoverySession(catalog,new SqlDocumentChunkStore(catalog),DocumentDecoderRegistry.CreateDefault());
        }
        public ISharePointCatalog Catalog { get { EnsureOpen(); return catalog; } }
        public RecoveryEngine Engine { get { EnsureOpen(); return engine; } }
        public DocumentExporter Exporter { get { EnsureOpen(); return exporter; } }
        public void Dispose()
        {
            if(disposed) return;
            disposed=true;
            engine.Close();
            try { var resource=store as IDisposable; if(resource!=null) resource.Dispose(); }
            finally { if(!Object.ReferenceEquals(store,catalog)) { var resource=catalog as IDisposable; if(resource!=null) resource.Dispose(); } }
        }
        private void EnsureOpen() { if(disposed) throw new ObjectDisposedException("RecoverySession"); }
    }
}