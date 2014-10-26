using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    // Summary:
    //     Specifies the way certain alerts and messages are handled while a macro is
    //     running.
    public enum WdAlertLevel
    {
        // Summary:
        //     Only message boxes are displayed; errors are trapped and returned to the
        //     macro.
        wdAlertsMessageBox = -2,
        //
        // Summary:
        //     All message boxes and alerts are displayed; errors are returned to the macro.
        wdAlertsAll = -1,
        //
        // Summary:
        //     No alerts or message boxes are displayed. If a macro encounters a message
        //     box, the default value is chosen and the macro continues.
        wdAlertsNone = 0,
    }
}
