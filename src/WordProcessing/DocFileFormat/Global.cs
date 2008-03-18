using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
	public class Global
	{
        public enum JustificationCode
        {
            left = 0,
            center,
            right,
            both,
            distribute,
            mediumKashida,
            numTab,
            highKashida,
            lowKashida,
            thaiDistribute,
        }

        public enum ColorIdentifier
        {
            auto = 0,
            black,
            blue,
            cyan,
            green,
            magenta,
            red,
            yellow,
            white,
            darkBlue,
            darkCyan,
            darkGreen,
            darkMagenta,
            darkRed,
            darkYellow,
            darkGray,
            lightGray,
        }

        public enum TextAnimation
        {
            none,
            lights,
            blinkBackground,
            sparkle,
            antsBlack,
            antsRed,
            shimmer
        }

        public enum UnderlineCode
        {
            none = 0,
            single,
            word,
            Double,
            dotted,
            notUsed1,
            thick,
            dash,
            notUsed2,
            dotDash,
            dotDotDash,
            wave,
            dottedHeavy,
            dashedHeavy,
            dashDotHeavy,
            dashDotDotHeavy,
            wavyHeavy,
            dashLong,
            wavyDouble,
            dashLongHeavy
        }

	}
}
