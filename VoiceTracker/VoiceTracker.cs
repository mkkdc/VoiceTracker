using System;
using System.ServiceProcess;
using System.Threading;
using System.IO;
using NAudio.Wave;
using NAudio.Lame;
using System.Speech.Synthesis;


namespace VoiceTracker
{
    public partial class VoiceTracker : ServiceBase{ // service class
        private string message; // initial message
        private StreamWriter file; // error log file
        private VoicerTracker application; // main application class
        private SpeechSynthesizer speech = new SpeechSynthesizer(); // synthesizer
        public VoiceTracker(){ // constructor
            InitializeComponent();
        }
        protected override void OnStart(string[] args){ // service start function 
            try{
                this.application = new VoicerTracker();
                eventLog1.WriteEntry("initial: service VoiceTracker is running..."); // system logs
                message = application.startProject(); //application start function
                eventLog1.WriteEntry(message); // system logs
            }
            catch (Exception error){ // ecxception error block 
                eventLog1.WriteEntry(error.Message);
                this.file = new StreamWriter(new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\errorLog.txt", System.IO.FileMode.Append));
                this.file.WriteLine(DateTime.Now + " - " + error.Message);
                this.file.Flush();
                this.file.Close();
                this.speech.Speak("Внимание произошла ошибка аудиозаписи,ошибка аудиозаписи.");
                this.Stop();
            }
        }
        protected override void OnStop(){ //stop function
            eventLog1.WriteEntry("initial: service VoiceTracker is stopped...");
            this.application.waveSource.StopRecording();
            VoicerTracker.lameWriter.Dispose();
        }
    }

    class VoicerTracker{ // main class
        private bool flag = false; // driving flag
        private DriveInfo[] drives; // drives array
        private DirectoryInfo dirInfo; // directory info
        public WaveInEvent waveSource = new WaveInEvent(); // record stream
        public static LameMP3FileWriter lameWriter; // mp3 file writer
        private static bool stopped = false; // writing flag 
        private DateTime dateStart = DateTime.Now; // date start time for end of today
        private DateTime dateEnd = DateTime.Now.Date.AddDays(1); // end date time for today
        private string tempFile; // file of recording 
        private string message;  // message return
        private int delay; // delay
        private StreamWriter file; // stream for error log file
        private SpeechSynthesizer speech = new SpeechSynthesizer();  // synthesizer
        private DriveInfo tempDrive; // temp drive
        private DirectoryInfo[] foldersArray; // folders for delete
        private Timer timer; // timer of recording

        public VoicerTracker(){} //constructor

        public string startProject(){  // main start function
              this.diskCheck(); // disk check function
              this.microphoneCheck(); //microphone check
              this.recording(); // record function
              return this.message;
        }

        private void diskCheck(){ // available space and drive cheacking function
             this.message += "initial: initialization of local volumes:\r\n";
             this.drives = DriveInfo.GetDrives();

             foreach (DriveInfo drive in drives) {
                 this.message += $"disk name: {drive.Name}\r\n";
                 this.message += $"disk type: {drive.DriveType}\r\n\r\n";
                 if ((drive.Name[0] == 'C' || drive.Name[0] == 'D') && System.Convert.ToString(drive.DriveType) == "Fixed"){
                    this.flag = true;
                    this.tempDrive = drive;
                 }
             }

             switch (flag){
                 case true:
                    this.message += "logical volume for recording discs detected.\r\n";
                    this.dirInfo = new DirectoryInfo(this.tempDrive.Name[0]+":\\VoiceTracker");
                    if (!this.dirInfo.Exists)
                        this.dirInfo.Create();
                        break;
                    case false:
                        throw new Exception("error: local write volume was found.");
             }

             if (this.tempDrive.AvailableFreeSpace < 1000000000)
                throw new Exception("error: insufficient free disk space.");

             this.foldersArray = this.dirInfo.GetDirectories();
             foreach (DirectoryInfo folder in foldersArray){
                 try{
                      if (Convert.ToDateTime(folder.Name) < this.dateStart.AddDays(-30))
                      folder.Delete();
                 }
                 catch{}
             }
        } 

        private void microphoneCheck(){ //microphone check function
             if (WaveIn.DeviceCount < 1)
                 throw new Exception("error: microphone was not detected");
             else
                 this.message += "initial: microphone detected\r\n";
        } 

        private void recording(){ //record function
            if (System.Convert.ToInt32(dateStart.ToString("mm")) <= 30)  
               this.delay = (29 - System.Convert.ToInt32(this.dateStart.ToString("mm"))) * 60 + (60 - System.Convert.ToInt32(this.dateStart.ToString("ss")));
            else
               this.delay = (59 - System.Convert.ToInt32(this.dateStart.ToString("mm"))) * 60 + (60 - System.Convert.ToInt32(this.dateStart.ToString("ss")));

            this.message += "initial: voicemail recording started...\r\n";

            this.waveSource = new WaveInEvent();
            this.waveSource.WaveFormat = new WaveFormat(44100,1); //hz,kbps,channels of record
            this.waveSource.DataAvailable += waveIn_DataAvailable;
            this.waveSource.RecordingStopped += waveIn_RecordingStopped;

            this.dirInfo = new DirectoryInfo(this.tempDrive.Name[0]+@":\VoiceTracker\" + this.dateStart.ToShortDateString());
            if(!this.dirInfo.Exists)
               this.dirInfo.Create();

            this.tempFile = (this.tempDrive.Name[0] + @":\VoiceTracker\" + this.dateStart.ToShortDateString() + @"\" + this.dateStart.ToShortDateString() + "-" + this.dateStart.ToString("HH.mm.ss") + @".mp3");
            lameWriter = new LameMP3FileWriter(this.tempFile, this.waveSource.WaveFormat,128);

            this.waveSource.StartRecording();
            stopped = false;
            this.timer = new Timer(_ => cicleRecording(), null, this.delay*1000, 1800*1000); // this.delay, 1800 <- interval ms
        } 

        static void waveIn_DataAvailable(object sender, WaveInEventArgs e){ //support function
            if(lameWriter != null)
               lameWriter.Write(e.Buffer, 0, e.BytesRecorded);
        }

        static void waveIn_RecordingStopped(object sender, StoppedEventArgs e){ //support function
            stopped = true;
        } 

        private void cicleRecording(){ // record cicle function
            this.waveSource.StopRecording();
            lameWriter.Flush();
            this.waveSource.Dispose();
            lameWriter.Dispose();

            this.fastCheck();

            this.dateStart = DateTime.Now;

            if (this.dateStart > this.dateEnd){
                this.terminte("completion of the application: The current day is over.");
            }

            this.waveSource = new WaveInEvent();
            this.waveSource.WaveFormat = new WaveFormat(44100,1); //hz,kbps,chanels of record
            this.waveSource.DataAvailable += waveIn_DataAvailable;
            this.waveSource.RecordingStopped += waveIn_RecordingStopped;

            this.dirInfo = new DirectoryInfo(this.tempDrive.Name[0] + @":\VoiceTracker\" + dateStart.ToShortDateString());
            if (!this.dirInfo.Exists)
               this.dirInfo.Create();
            this.tempFile = (this.tempDrive.Name[0] + @":\VoiceTracker\" + dateStart.ToShortDateString() + @"\" + dateStart.ToShortDateString() + "-" + dateStart.ToString("HH.mm.ss") + @".mp3");

            lameWriter = new LameMP3FileWriter(this.tempFile, this.waveSource.WaveFormat,128);
            this.waveSource.StartRecording();     
        } 

        private void fastCheck(){ //fast check function
            this.drives = DriveInfo.GetDrives();
            foreach (DriveInfo drive in drives){
               if (drive.Name[0] == this.tempDrive.Name[0]  && System.Convert.ToString(drive.DriveType) == "Fixed"){
                  flag = true;
                  if (drive.AvailableFreeSpace < 500000000){
                     this.terminte("error: fastCheck failed - not enought space.");
                  }
               }
            }

            switch (flag){
               case true:
                   this.dirInfo = new DirectoryInfo(this.tempDrive.Name[0] + ":\\VoiceTracker");
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

        private void terminte(string message){ //terminate function
             this.file = new StreamWriter(new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\errorLog.txt", System.IO.FileMode.Append));
             this.file.WriteLine(DateTime.Now + " - " + message);
             this.file.Flush();
             this.file.Close();
             this.speech.Speak("Внимание произошла ошибка аудиозаписи,ошибка аудиозаписи.");
             Environment.Exit(1);
        } 
    }
}
