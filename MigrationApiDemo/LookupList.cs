using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrationApiDemo
{
    public class LookupList
    {
        public string listId { get; set; }
        public ListItemCollection itemArray { get; set; }
    }
}
