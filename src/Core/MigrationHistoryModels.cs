using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
namespace SharePointExplorer
{
    public interface ISharePointMigrationHistoryCatalog
    {
        MigrationHistorySnapshot ReadMigrationHistory(MigrationListSnapshot snapshot);
    }
    public sealed class MigrationHistorySnapshot
    {
        public ReadOnlyCollection<MigrationItemSnapshot> Items {get;private set;}
        public ReadOnlyCollection<MigrationUserSnapshot> Users {get;private set;}
        public ReadOnlyCollection<MigrationHistoryFile> Files {get;private set;}
        public MigrationHistorySnapshot(IEnumerable<MigrationItemSnapshot> items,IEnumerable<MigrationUserSnapshot> users,IEnumerable<MigrationHistoryFile> files)
        {Items=new List<MigrationItemSnapshot>(items??Enumerable.Empty<MigrationItemSnapshot>()).AsReadOnly();Users=new List<MigrationUserSnapshot>(users??Enumerable.Empty<MigrationUserSnapshot>()).AsReadOnly();Files=new List<MigrationHistoryFile>(files??Enumerable.Empty<MigrationHistoryFile>()).AsReadOnly();}
        public static MigrationHistorySnapshot Empty {get{return new MigrationHistorySnapshot(null,null,null);}}
    }
    public sealed class MigrationHistoryFile
    {
        private readonly Node document;
        public Node Document {get{return MigrationSnapshotCopy.Node(document);}}
        public string CheckinComment {get;private set;}
        public MigrationHistoryFile(Node document,string checkinComment)
        {if(document==null)throw new ArgumentNullException("document");this.document=MigrationSnapshotCopy.Node(document);CheckinComment=checkinComment;}
    }
}
