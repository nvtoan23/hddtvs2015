namespace dllbhyt
{
    using System;
    using System.Diagnostics;

    public class run
    {
        private string arguments;
        private string fileName;
        private int waitMillisecs;
        private bool waitUntilFinished;

        public run()
        {
            this.fileName = string.Empty;
            this.arguments = string.Empty;
            this.waitUntilFinished = false;
            this.waitMillisecs = -1;
        }

        public run(string fileName)
        {
            this.fileName = string.Empty;
            this.arguments = string.Empty;
            this.waitUntilFinished = false;
            this.waitMillisecs = -1;
            this.fileName = fileName;
        }

        public run(string fileName, string arguments, bool waitUntilFinished)
        {
            this.fileName = string.Empty;
            this.arguments = string.Empty;
            this.waitUntilFinished = false;
            this.waitMillisecs = -1;
            this.fileName = fileName;
            this.arguments = arguments;
            this.waitUntilFinished = waitUntilFinished;
        }

        public run(string fileName, string arguments, int waitMillisecs)
        {
            this.fileName = string.Empty;
            this.arguments = string.Empty;
            this.waitUntilFinished = false;
            this.waitMillisecs = -1;
            this.fileName = fileName;
            this.arguments = arguments;
            this.waitMillisecs = waitMillisecs;
        }

        public void Launch()
        {
            using (Process process = new Process())
            {
                process.StartInfo.FileName = this.fileName;
                if (this.arguments.Length > 0)
                {
                    process.StartInfo.Arguments = this.arguments;
                }
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                if (this.waitMillisecs != -1)
                {
                    process.WaitForExit(this.waitMillisecs);
                }
                else if (this.waitUntilFinished)
                {
                    process.WaitForExit();
                }
                process.Close();
            }
        }
    }
}

