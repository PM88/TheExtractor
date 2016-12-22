using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ADOX;

namespace Report_generator
{
    public class AdodbCommandDraft
    {
        public ADODB.Command Command;
        //public ADODB.Connection ActiveConnection;

        public AdodbCommandDraft(ADODB.Connection connection) { Command = new ADODB.Command(); Command.ActiveConnection = connection; }


    }
}
