using OutlookExchangeOnlineAPI;
using OutlookSql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Messages
{

    public class ParseOutlookItem
    {
        public Event Item { get; protected set; }

        public ParseOutlookItem(Event item)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
        }
    }
}
