using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Patient_Master
{
    
    public class LogWriter
    {
        private string m_exePath = string.Empty;
        public void LogWrite(string logMessage)
        {
            m_exePath = Path.GetTempPath();
            try
            {
                string Logname = "Patent_log_" + DateTime.Now.Date.ToString("ddMMyyyy") + ".txt";
                using (StreamWriter w = File.AppendText(Path.Combine(m_exePath, Logname)))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
                txtWriter.WriteLine("  :");
                txtWriter.WriteLine("  :{0}", logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
            }
        }
    }
}
