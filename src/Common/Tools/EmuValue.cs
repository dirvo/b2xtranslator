using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.Tools
{

    public class EmuValue
    {
        private int value;

        /// <summary>
        /// Creates a new EmuValue for the given value.
        /// </summary>
        /// <param name="value"></param>
        public EmuValue(int value)
        {
            this.value = value;
        }

        /// <summary>
        /// Converts the EMU to pt
        /// </summary>
        /// <returns></returns>
        public double ToPoints()
        {
            return this.value / 12700;
        }

        /// <summary>
        /// Converts the EMU to twips
        /// </summary>
        /// <returns></returns>
        public double ToTwips()
        {
            return this.value / 635;
        }
    }
}
