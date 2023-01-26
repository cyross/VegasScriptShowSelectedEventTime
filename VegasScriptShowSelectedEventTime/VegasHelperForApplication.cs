using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VegasScriptShowSelectedEventTime
{
    internal partial class VegasHelper
    {
        internal void AddTrackEventStateChangedEventHandler(EventHandler handler)
        {
            Vegas.TrackEventStateChanged += handler;
        }

    }
}
