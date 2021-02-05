using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ERPInquire.Base
{
    class Log
    {
        public long Id { get; set; }
        /// <summary>
        /// 日志级别 Trace|Debug|Info|Warn|Error|Fatal
        /// </summary>
        public string Level { get; set; }
        public string Message { get; set; }
        public string Action { get; set; }
        public string usersname { get; set; }
        public DateTime Timestamp { get; set; }

        private Log() { }
        public Log(string level, string message, string action = null, string usersname = null)
        {
            this.Level = level;
            this.Message = message;
            this.Action = action;
            this.usersname = usersname;
        }
    }
}