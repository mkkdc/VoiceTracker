using System;
using System.ServiceProcess;
using System.Threading.Tasks;
using System.IO;
using NAudio.Wave;


namespace VoiceTracker
{
    public partial class VoiceTracker : ServiceBase
    {
        private string message;
        private StreamWriter file;
        VoicerTracker application;
        public VoiceTracker(){
            InitializeComponent();
        }
        protected override void OnStart(string[] args){
            try{
                this.application = new VoicerTracker();
                eventLog1.WriteEntry("initial: service VoiceTracker is running...");
                message = application.startProject();
                eventLog1.WriteEntry(message);
            }
            catch (Exception error){
                eventLog1.WriteEntry(error.Message);
                this.file = new StreamWriter(new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\errorLog.txt", System.IO.FileMode.Append));
                this.file.WriteLine(DateTime.Now + " - " + error.Message);
                this.file.Flush();
                this.file.Close();
                this.Stop();
            }
        }
        protected override void OnStop(){
            eventLog1.WriteEntry("initial: service VoiceTracker is stopped...");
            this.application.waveSource.StopRecording();
            this.application.waveFile.Dispose();
        }
    }

    class VoicerTracker
    {
        private bool flag = false;
        private DriveInfo[] drives;
        private DirectoryInfo dirInfo = new DirectoryInfo("D:\\VoiceTracker");
        public WaveFileWriter waveFile;
        public WaveInEvent waveSource = new WaveInEvent();
        private DateTime dateStart = DateTime.Now;
        private DateTime dateEnd = DateTime.Now.Date.AddDays(1);
        private string tempFile;
        private string message;
        private int delay;
        private StreamWriter file;

        public VoicerTracker(){}

        public string startProject(){
            this.diskCheck();
            this.microphoneCheck();
            this.recording();
            return this.message;
        }

        private void diskCheck(){
            this.message += "initial: initialization of local volumes:\r\n";
            this.drives = DriveInfo.GetDrives();
   
            foreach (DriveInfo drive in drives){
                this.message += $"disk name: {drive.Name}\r\n";
                this.message += $"disk type: {drive.DriveType}\r\n\r\n";
                if (drive.Name[0] == 'D' && System.Convert.ToString(drive.DriveType) == "Fixed"){
                    flag = true;
                    if (drive.AvailableFreeSpace < 1000000000)
                        throw new Exception("error: insufficient free disk space.");
                }
            }

            switch (flag){
                case true:
                    this.message += "logical volume for recording discs detected.\r\n";
                    if (!this.dirInfo.Exists)
                        this.dirInfo.Create();
                    break;
                case false:
                    throw new Exception("error: local write volume was found.");
            }
        }

        private void microphoneCheck(){
            if (WaveIn.DeviceCount < 1)
                throw new Exception("error: microphone was not detected");
            else
                this.message += "initial: microphone detected\r\n";
        }

        private void recording(){
            if (System.Convert.ToInt32(dateStart.ToString("mm")) <= 30)
                this.delay = (29 - System.Convert.ToInt32(this.dateStart.ToString("mm"))) * 60 + (60 - System.Convert.ToInt32(this.dateStart.ToString("ss")));
            else
                this.delay = (59 - System.Convert.ToInt32(this.dateStart.ToString("mm"))) * 60 + (60 - System.Convert.ToInt32(this.dateStart.ToString("ss")));

            this.message += "initial: voicemail recording started...\r\n";

            this.waveSource.DeviceNumber = 0;
            this.waveSource.WaveFormat = new WaveFormat(8000, 1);
            this.waveSource.DataAvailable += new EventHandler<WaveInEventArgs>(this.waveSource_DataAvailable);
            this.dirInfo = new DirectoryInfo(@"D:\VoiceTracker\" + this.dateStart.ToShortDateString());
            if (!this.dirInfo.Exists)
                this.dirInfo.Create();

            this.tempFile = (@"D:\VoiceTracker\" + this.dateStart.ToShortDateString() + @"\" + this.dateStart.ToShortDateString() + "-" + this.dateStart.ToString("HH.mm.ss") + @".wav");
            this.waveFile = new WaveFileWriter(tempFile, waveSource.WaveFormat);
            this.waveSource.StartRecording();

            SetInterval(() => cicleRecording(), TimeSpan.FromSeconds(5)); //this.delay
        }

        private void waveSource_DataAvailable(object sender, WaveInEventArgs e){
            waveFile.WriteData(e.Buffer, 0, e.BytesRecorded);
        }

        public static async Task SetInterval(Action action, TimeSpan timeout){
            await Task.Delay(timeout).ConfigureAwait(false);
            action();
            SetInterval(action, TimeSpan.FromSeconds(5));//1800

        }

        private void cicleRecording(){
            this.waveSource.StopRecording();
            this.waveFile.Dispose();
            this.fastCheck();
            this.dateStart = DateTime.Now;
            
            if (this.dateStart > this.dateEnd){
                this.terminte("completion of the application: The current day is over.");
            }

            this.waveSource = new WaveInEvent();
            this.waveSource.DeviceNumber = 0;
            this.waveSource.WaveFormat = new WaveFormat(8000, 1);
            this.waveSource.DataAvailable += new EventHandler<WaveInEventArgs>(waveSource_DataAvailable);
            this.dirInfo = new DirectoryInfo(@"D:\VoiceTracker\" + dateStart.ToShortDateString());
            if (!this.dirInfo.Exists)
                this.dirInfo.Create();
            this.tempFile = (@"D:\VoiceTracker\" + dateStart.ToShortDateString() + @"\" + dateStart.ToShortDateString() + "-" + dateStart.ToString("HH.mm.ss") + @".wav");
            Console.WriteLine(@"D:\VoiceTracker\" + dateStart.ToShortDateString() + @"\" + dateStart.ToShortDateString() + "-" + dateStart.ToString("HH.mm.ss") + @".wav");
            this.waveFile = new WaveFileWriter(tempFile, waveSource.WaveFormat);
            this.waveSource.StartRecording();
        }

        private void fastCheck(){
            this.drives = DriveInfo.GetDrives();
            foreach (DriveInfo drive in drives){
                if (drive.Name[0] == 'D' && System.Convert.ToString(drive.DriveType) == "Fixed"){
                    flag = true;
                    if (drive.AvailableFreeSpace < 500000000){
                        this.terminte("error: fastCheck failed - not enought space.");
                    }
                }
            }
            switch (flag){
                case true:
                    if (!this.dirInfo.Exists)
                        this.dirInfo.Create();
                    break;
                case false:
                    this.terminte("error: fastCheck failed - logical disk not found.");
                    break;

            }

            if (WaveIn.DeviceCount < 1){
                this.terminte("error: fastCheck failed - microphone is not found.");
            }
        }

        private void terminte(string message){
            this.file = new StreamWriter(new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\errorLog.txt", System.IO.FileMode.Append));
            this.file.WriteLine(DateTime.Now + " - " + message);
            this.file.Flush();
            this.file.Close();
            Environment.Exit(1);
        }
    }
}
