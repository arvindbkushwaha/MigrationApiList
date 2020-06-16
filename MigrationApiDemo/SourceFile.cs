using System;
using System.Collections.Generic;

namespace MigrationApiDemo
{
    public class SourceFile
    {
        public SourceFile()
        {
            Properties = new Dictionary<string, string>();
        }

        public DateTime LastModified { get; set; }
        public string Title { get; set; }
        public Dictionary<string,string> Properties { get; set; }
    }
}