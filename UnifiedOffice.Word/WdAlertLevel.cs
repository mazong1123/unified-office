using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnifiedOffice.Word
{
    /// <summary>
    /// Specifies the way certain alerts and messages are handled while a macro is
    /// running.
    /// </summary>
    public enum WdAlertLevel
    {
        /// <summary>
        /// Only message boxes are displayed; errors are trapped and returned to the
        /// macro.
        /// </summary>
        wdAlertsMessageBox = -2,

        /// <summary>
        /// All message boxes and alerts are displayed; errors are returned to the macro.
        /// </summary>
        wdAlertsAll = -1,
        //
        // Summary:
        //     No alerts or message boxes are displayed. If a macro encounters a message
        //     box, the default value is chosen and the macro continues.

        /// <summary>
        /// No alerts or message boxes are displayed. If a macro encounters a message
        /// box, the default value is chosen and the macro continues.
        /// </summary>
        wdAlertsNone = 0,
    }
}
