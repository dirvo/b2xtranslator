/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */


using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.Utils
{
    public static class TraceLogger
    {

        public enum LoggingLevel
        {
            DEBUG = 0,
            INFO,
            WARNING,
            ERROR
        }


        static LoggingLevel _logLevel = LoggingLevel.WARNING;
        public static LoggingLevel LogLevel
        {
            get { return TraceLogger._logLevel; }
            set { TraceLogger._logLevel = value; }
        }


        private static void WriteLine(string msg, LoggingLevel level)
        {
            if (_logLevel <= level)
                System.Diagnostics.Trace.WriteLine(string.Format("{0} " + msg, System.DateTime.Now));
        }


        public static void Debug(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[D] " + string.Format(msg, objs), LoggingLevel.DEBUG);
        }


        public static void Info(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[I] " + string.Format(msg, objs), LoggingLevel.INFO);
        }


        public static void Warning(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[W] " + string.Format(msg, objs), LoggingLevel.WARNING);
        }


        public static void Error(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[E] " + string.Format(msg, objs), LoggingLevel.ERROR);
        }
    }
}

