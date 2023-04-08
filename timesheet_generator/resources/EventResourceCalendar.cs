using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace timesheet_generator.resources
{

    public enum EventTypeCalendar
    {
        Work,
        WorkRemote
    }
    public sealed record EventResourceCalendar(string Title, EventTypeCalendar eventType);
}
